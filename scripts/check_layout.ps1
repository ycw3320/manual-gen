<#
.SYNOPSIS
표가 있는 쪽의 실제 렌더 겹침 검사 — 빌더(build_pptx.py)가 표 높이를 글꼴 실측으로 잡았을 때
결과를 PowerPoint 로 확인하는 게이트.

.DESCRIPTION
python-pptx 가 파일에 적는 표 높이는 요청값일 뿐, 실제 행 높이는 PowerPoint 가 글자를 배치하며
정한다. 이 스크립트는 산출물을 읽기 전용으로 열어, 표가 있는 쪽마다
  - 표의 실제 아래 끝이 그 아래 도형(설명 글·그림)의 위 끝을 넘는지(겹침)
  - 표의 실제 아래 끝이 본문 하한을 넘는지
  - 그 쪽 글상자의 실제 글자 끝(BoundTop+BoundHeight)이 본문 하한을 넘는지
를 잰다. 머리·꼬리·장식(이름표 mg-chrome·mg-art)과 빈 글상자는 제외한다.

출력: 문제마다 "VIOLATION slide=<n> what=<table_overlap|table_bottom|text_bottom> over_in=<넘친 길이>"
한 줄(콘솔 인코딩과 무관하게 읽히도록 영문), 마지막에 요약 한 줄.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts/check_layout.ps1 -Path out.pptx -BottomIn 10.3

종료 코드: 0 문제 없음 / 1 겹침·넘침 있음 / 2 PowerPoint 로 열 수 없음
#>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [double]$BottomIn = 10.3,
    [double]$TolIn = 0.03
)
$ErrorActionPreference = "Stop"
try {
    $pp = New-Object -ComObject PowerPoint.Application
} catch {
    Write-Output "[check_layout] PowerPoint 를 열 수 없습니다"
    exit 2
}
$bad = 0
$tables = 0
try {
    try {
        $pres = $pp.Presentations.Open((Resolve-Path $Path).Path, $true, $false, $false)
    } catch {
        Write-Output "[check_layout] 파일을 열 수 없습니다: $Path"
        exit 2
    }
    $limit = ($BottomIn + $TolIn) * 72
    foreach ($s in $pres.Slides) {
        $body = @()
        foreach ($sh in $s.Shapes) {
            if ($sh.Name -like "mg-chrome*" -or $sh.Name -like "mg-art*") { continue }
            $isTable = ($sh.HasTable -eq -1)
            $hasText = (-not $isTable) -and ($sh.HasTextFrame -eq -1) -and ($sh.TextFrame.HasText -eq -1)
            $isPic = ($sh.Type -eq 13)
            if (-not ($isTable -or $hasText -or $isPic)) { continue }
            $bottom = $sh.Top + $sh.Height
            if ($hasText) {
                $tr = $sh.TextFrame.TextRange
                $bottom = $tr.BoundTop + $tr.BoundHeight
            }
            $body += [pscustomobject]@{ Table = $isTable; Text = $hasText; Top = $sh.Top; Bottom = $bottom }
        }
        $tbls = @($body | Where-Object { $_.Table })
        if ($tbls.Count -eq 0) { continue }
        $tables += $tbls.Count
        foreach ($t in $tbls) {
            $below = @($body | Where-Object { -not $_.Table -and $_.Top -gt $t.Top + 1 } | ForEach-Object { $_.Top })
            $below += @($tbls | Where-Object { $_.Top -gt $t.Top + 1 } | ForEach-Object { $_.Top })
            if ($below.Count -gt 0) {
                $next = ($below | Measure-Object -Minimum).Minimum
                if ($t.Bottom -gt $next - 1.5) {
                    Write-Output ("VIOLATION slide={0} what=table_overlap over_in={1:N3}" -f $s.SlideIndex, (($t.Bottom - $next + 1.5) / 72))
                    $bad++
                }
            }
            if ($t.Bottom -gt $limit) {
                Write-Output ("VIOLATION slide={0} what=table_bottom over_in={1:N3}" -f $s.SlideIndex, (($t.Bottom - $limit) / 72))
                $bad++
            }
        }
        foreach ($x in @($body | Where-Object { $_.Text })) {
            if ($x.Bottom -gt $limit) {
                Write-Output ("VIOLATION slide={0} what=text_bottom over_in={1:N3}" -f $s.SlideIndex, (($x.Bottom - $limit) / 72))
                $bad++
            }
        }
    }
    $pres.Close()
} finally {
    if ($pp) { $pp.Quit() }
}
Write-Output ("[check_layout] 표 {0}개 검사 — 겹침·넘침 {1}건" -f $tables, $bad)
exit $(if ($bad -gt 0) { 1 } else { 0 })
