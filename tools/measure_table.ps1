<#
.SYNOPSIS
표 기하 보정 2단계 — PowerPoint 가 실제로 렌더한 행 높이를 잰다.

.DESCRIPTION
python-pptx 가 파일에 적는 표 높이는 '요청값'이고, 실제 행 높이는 PowerPoint 가
텍스트를 배치하면서 다시 정한다. 따라서 보정의 기준값은 COM 으로 열어서 읽어야 한다.
문서를 ReadOnly 로 열므로 같은 파일을 열어 둔 상태여도 동작한다.

산출 CSV 열: slide, shape_top_pt, shape_h_pt, row_idx, row_h_pt  (단위 pt, 1in = 72pt)

.EXAMPLE
powershell -ExecutionPolicy Bypass -File tools/measure_table.ps1 `
  -Path C:\work\probe.pptx -Out C:\work\measured.csv

종료 코드: 0 성공 / 1 실패(PowerPoint 없음 포함)
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = "Stop"

try {
  $pp = New-Object -ComObject PowerPoint.Application
} catch {
  Write-Error "[measure] PowerPoint 를 열 수 없습니다 — 이 보정은 PowerPoint COM 이 필요합니다."
  exit 1
}

try {
  $pres = $pp.Presentations.Open($Path, $true, $false, $false)   # ReadOnly, 창 없이
  $rows = New-Object System.Collections.Generic.List[string]
  $rows.Add("slide,shape_top_pt,shape_h_pt,row_idx,row_h_pt")
  foreach ($s in $pres.Slides) {
    foreach ($sh in $s.Shapes) {
      if ($sh.HasTable -eq -1) {
        $t = $sh.Table
        for ($r = 1; $r -le $t.Rows.Count; $r++) {
          $rows.Add("$($s.SlideIndex),$($sh.Top),$($sh.Height),$r,$($t.Rows.Item($r).Height)")
        }
      }
    }
  }
  $pres.Close()
  [System.IO.File]::WriteAllLines($Out, $rows)
  Write-Output "[measure] 행 $($rows.Count - 1)건 측정 → $Out"
} finally {
  if ($pp) { $pp.Quit() }
}
