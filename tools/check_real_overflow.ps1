<#
.SYNOPSIS
산출물이 실제로 렌더됐을 때 페이지를 넘는 도형이 있는지 검사한다.

.DESCRIPTION
verify_pptx.py 는 파일에 적힌 좌표를 읽는다. 그런데 표는 PowerPoint 가 텍스트를
배치하면서 행 높이를 다시 정하므로, 파일에 적힌 값과 실제 렌더 결과가 다르다.
표 기하 계수를 손댄 뒤에는 이 검사로 실제 렌더 기준 넘침이 0 인지 확인해야 한다.

일상 검증은 verify_pptx.py 로 충분하다 — 이 스크립트는 계수를 바꿨을 때만 쓴다.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File tools/check_real_overflow.ps1 `
  -Path C:\work\관리자매뉴얼_시스템명_20260811.pptx

종료 코드: 0 넘침 없음 / 1 넘침 있음 / 2 PowerPoint 불가
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [double]$TolerancePt = 0.5      # 반올림 오차 흡수
)
$ErrorActionPreference = "Stop"

try {
  $pp = New-Object -ComObject PowerPoint.Application
} catch {
  Write-Error "[real] PowerPoint 를 열 수 없습니다 — 이 검사는 PowerPoint COM 이 필요합니다."
  exit 2
}

$bad = 0
try {
  try {
    $pres = $pp.Presentations.Open($Path, $true, $false, $false)   # ReadOnly, 창 없이
  } catch {
    # 파일을 못 여는 것과 '넘침 있음'은 다른 상황이다 — 종료 코드로 구분한다.
    # ErrorActionPreference=Stop 에서 Write-Error 는 종료성이라 exit 2 에 닿지 못한다.
    $host.UI.WriteErrorLine("[real] 파일을 열 수 없습니다: $Path")
    $host.UI.WriteErrorLine("  $($_.Exception.Message)")
    exit 2
  }
  $H = $pres.PageSetup.SlideHeight
  foreach ($s in $pres.Slides) {
    foreach ($sh in $s.Shapes) {
      $bottom = $sh.Top + $sh.Height
      if ($bottom -gt ($H + $TolerancePt)) {
        $kind = if ($sh.HasTable -eq -1) { "표" } else { "도형" }
        Write-Output ("  슬라이드 {0} {1}: 아래로 {2:N2}in 넘침" -f $s.SlideIndex, $kind, (($bottom - $H) / 72))
        $bad++
      }
    }
  }
  $pres.Close()
  Write-Output ("[real] {0} — 실제 렌더 넘침 {1}건" -f (Split-Path $Path -Leaf), $bad)
} finally {
  if ($pp) { $pp.Quit() }
}

exit $(if ($bad -gt 0) { 1 } else { 0 })
