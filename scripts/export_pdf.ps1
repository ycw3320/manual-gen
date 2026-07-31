<#
.SYNOPSIS
매뉴얼 pptx → PDF 변환 (최종 산출물용).

.DESCRIPTION
이 skill 의 산출 형식은 pptx 와 PDF 두 가지이며, PDF 는 **완성된 pptx 를 변환해서**
만든다 — 레이아웃·규격(고정 프레임·테두리·타일 분할·캡션)이 pptx 빌더에 내장되어
있으므로 PDF 를 따로 조판하지 않는다. 따라서 검증(verify_pptx)도 pptx 단계에서 끝난다.

PowerPoint COM 의 ExportAsFixedFormat 을 우선 사용하고(글꼴·도형 재현이 가장 정확),
PowerPoint 가 없으면 LibreOffice(soffice) 변환으로 폴백한다. 문서를 ReadOnly 로 열므로
사용자가 같은 파일을 열어 둔 상태여도 동작한다.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts/export_pdf.ps1 `
  -Path 관리자매뉴얼_시스템명_20260724.pptx

산출: 같은 폴더에 동일 이름의 .pdf (또는 -Out 으로 지정)

종료 코드: 0 성공 / 1 실패(PowerPoint·soffice 모두 불가 포함)
#>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Out
)

$ErrorActionPreference = "Stop"

$resolved = Resolve-Path $Path -ErrorAction SilentlyContinue
if ($null -eq $resolved) {
    Write-Error "[export_pdf] 파일 없음: $Path"
    exit 1
}
$abs = $resolved.Path
if (-not $Out) {
    $Out = [System.IO.Path]::ChangeExtension($abs, ".pdf")
}
$outDir = Split-Path $Out
if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Force $outDir | Out-Null
}

# 1) PowerPoint COM — ExportAsFixedFormat (표준 경로)
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
        # SaveAs(FileName, FileFormat) — 32 = ppSaveAsPDF
        # (ExportAsFixedFormat 은 PowerShell 에서 인자 타입 변환이 실패한다)
        $pres.SaveAs($Out, 32)
        $pres.Close()
        Write-Output "[export_pdf] PDF 저장 완료: $Out"
        exit 0
    } catch {
        Write-Error "[export_pdf] PowerPoint 변환 실패: $_"
        exit 1
    } finally {
        $pp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pp) | Out-Null
    }
}

# 2) LibreOffice 폴백
$soffice = Get-Command soffice -ErrorAction SilentlyContinue
if ($null -eq $soffice) {
    foreach ($cand in @("C:\Program Files\LibreOffice\program\soffice.exe",
                        "C:\Program Files (x86)\LibreOffice\program\soffice.exe")) {
        if (Test-Path $cand) { $soffice = @{ Source = $cand }; break }
    }
}
if ($null -ne $soffice) {
    $tmpDir = Split-Path $Out
    if (-not $tmpDir) { $tmpDir = "." }
    & $soffice.Source --headless --convert-to pdf --outdir $tmpDir $abs | Out-Null
    $made = Join-Path $tmpDir ([System.IO.Path]::GetFileNameWithoutExtension($abs) + ".pdf")
    if (Test-Path $made) {
        if ($made -ne $Out) { Move-Item -Force $made $Out }
        Write-Output "[export_pdf] PDF 저장 완료(LibreOffice): $Out"
        exit 0
    }
}

Write-Error "[export_pdf] PowerPoint·LibreOffice 모두 사용 불가 — pptx 를 열어 '다른 이름으로 저장 > PDF' 로 변환하세요"
exit 1
