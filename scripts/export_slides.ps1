<#
.SYNOPSIS
pptx 렌더 검수용 슬라이드 PNG export.

.DESCRIPTION
생성된 pptx 를 슬라이드별 PNG 로 렌더해 시각 검수(이미지 누락/겹침/캡션 대응/한글
깨짐)를 가능하게 한다. PowerPoint COM(사내 표준 환경) 우선이며, PowerPoint 가 없으면
LibreOffice(soffice) PDF 변환으로 폴백한다(PDF 는 뷰어에서 페이지 넘겨 검수).
문서를 ReadOnly 로 열므로 사용자가 같은 파일을 열어 둔 상태여도 동작한다.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts/export_slides.ps1 `
  -Path 관리자매뉴얼_시스템명_20260714.pptx -OutDir manual-work\render-check

종료 코드: 0 성공 / 1 실패(PowerPoint·soffice 모두 불가 포함)
#>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$OutDir,
    [int]$Width = 1200
)

$ErrorActionPreference = "Stop"

$resolved = Resolve-Path $Path -ErrorAction SilentlyContinue
if ($null -eq $resolved) {
    Write-Error "[export_slides] 파일 없음: $Path"
    exit 1
}
$abs = $resolved.Path
if (-not $OutDir) {
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($abs)
    $OutDir = Join-Path (Split-Path $abs) ("render-check\" + $stem)
}
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = [System.IO.Path]::GetFullPath($OutDir)

# 1) PowerPoint COM — 슬라이드별 PNG (표준 경로)
$pp = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
} catch {
    $pp = $null
}
if ($null -ne $pp) {
    try {
        # Open(FileName, ReadOnly, Untitled, WithWindow)
        $pres = $pp.Presentations.Open($abs, $true, $false, $false)
        $total = $pres.Slides.Count
        foreach ($slide in $pres.Slides) {
            $png = Join-Path $OutDir ("s{0:d2}.png" -f $slide.SlideIndex)
            $slide.Export($png, "PNG", $Width, 0)
        }
        $pres.Close()
        Write-Output "[export_slides] PNG $total 장 저장: $OutDir"
        exit 0
    } catch {
        Write-Error "[export_slides] PowerPoint export 실패: $_"
        exit 1
    } finally {
        $pp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pp) | Out-Null
    }
}

# 2) LibreOffice 폴백 — PDF 변환 (soffice 의 PNG 변환은 첫 슬라이드만 출력하므로 PDF)
$soffice = Get-Command soffice -ErrorAction SilentlyContinue
if ($null -eq $soffice) {
    foreach ($cand in @("C:\Program Files\LibreOffice\program\soffice.exe",
                        "C:\Program Files (x86)\LibreOffice\program\soffice.exe")) {
        if (Test-Path $cand) { $soffice = @{ Source = $cand }; break }
    }
}
if ($null -ne $soffice) {
    & $soffice.Source --headless --convert-to pdf --outdir $OutDir $abs | Out-Null
    $pdf = Join-Path $OutDir ([System.IO.Path]::GetFileNameWithoutExtension($abs) + ".pdf")
    if (Test-Path $pdf) {
        Write-Output "[export_slides] PowerPoint 미설치 — PDF 로 변환: $pdf (뷰어로 페이지별 검수)"
        exit 0
    }
}

Write-Error "[export_slides] PowerPoint COM·LibreOffice 모두 사용 불가 — 렌더 검수를 건너뛰려면 pptx 를 직접 열어 육안 확인하세요"
exit 1
