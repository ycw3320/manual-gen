<#
.SYNOPSIS
파이썬 확인(필요하면 설치) — skill 의 다른 스크립트가 모두 파이썬이라, 파이썬 없이 도는
이 스크립트로 맨 먼저 확인한다.

.DESCRIPTION
쓸 수 있는 Python 3.10 이상을 찾아 경로를 알려 준다. 찾는 순서:
  py 런처(-3) → PATH 의 python / python3 → 기본 설치 경로(사용자·전체 사용자)
Windows 의 WindowsApps\python.exe 는 대개 Microsoft Store 로 연결되는 바로가기다(설치 아님).
인자 없이 실행하면 Store 창이 열리므로 실행하지 않고, Store 판 Python 이 실제로 설치돼
있을 때만 후보로 쓴다.

-Install 은 사용자가 설치에 동의한 뒤에만 쓴다: winget 으로 Python 3.12 를 **현재 사용자에게만**
설치한다(관리자 권한 불필요, Microsoft 공식 winget 원본). 설치하면 Python 라이선스(PSF)와
winget 원본 약관에 동의하게 된다. 설치 직후에는 이 창의 PATH 에 반영되지 않으므로 출력된
경로로 이어서 실행한다.

출력 첫 두 줄(기계 판독용):
  PYTHON=<python.exe 경로 | 빈 값>
  VERSION=<3.x.y | 빈 값>

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts/check_python.ps1
powershell -ExecutionPolicy Bypass -File scripts/check_python.ps1 -Install

종료 코드: 0 사용 가능 / 1 없음(설치 필요) / 2 3.10 미만만 있음 / 3 설치 실패·winget 없음
#>
param([switch]$Install)

$MinMajor, $MinMinor = 3, 10
$ErrorActionPreference = "SilentlyContinue"

function Probe([string]$exe, [string[]]$pre = @()) {
    # 실행해 보고 (경로, 버전)을 돌려준다 — 실패하면 $null
    $out = & $exe @pre -c "import sys;print(sys.executable);print('%d.%d.%d' % sys.version_info[:3])" 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $out -or @($out).Count -lt 2) { return $null }
    return [pscustomobject]@{ Path = @($out)[0].Trim(); Version = @($out)[1].Trim() }
}

function IsStoreStub([string]$path) {
    # WindowsApps 의 바로가기 — Store 판 Python 패키지가 없으면 설치본이 아니다
    if ($path -notlike "*\WindowsApps\*") { return $false }
    $pkg = Get-ChildItem "$env:LOCALAPPDATA\Packages" -Directory -Filter "PythonSoftwareFoundation.Python.3*"
    return -not $pkg
}

function Find-Python {
    $cands = New-Object System.Collections.Generic.List[object]
    $py = Get-Command py -CommandType Application | Where-Object { -not (IsStoreStub $_.Source) } | Select-Object -First 1
    if ($py) { $cands.Add(@($py.Source, @("-3"))) }
    foreach ($name in "python", "python3") {
        foreach ($c in @(Get-Command $name -All -CommandType Application)) {
            if ($c -and -not (IsStoreStub $c.Source)) { $cands.Add(@($c.Source, @())) }
        }
    }
    $roots = @("$env:LOCALAPPDATA\Programs\Python", "$env:ProgramFiles", "${env:ProgramFiles(x86)}", "C:\")
    foreach ($r in $roots) {
        Get-ChildItem $r -Directory -Filter "Python3*" | Sort-Object Name -Descending | ForEach-Object {
            $exe = Join-Path $_.FullName "python.exe"
            if (Test-Path $exe) { $cands.Add(@($exe, @())) }
        }
    }
    $old = $null
    foreach ($c in $cands) {
        $r = Probe $c[0] $c[1]
        if (-not $r) { continue }
        $v = $r.Version.Split(".")
        if ([int]$v[0] -gt $MinMajor -or ([int]$v[0] -eq $MinMajor -and [int]$v[1] -ge $MinMinor)) { return $r }
        if (-not $old) { $old = $r }
    }
    if ($old) { $old | Add-Member -NotePropertyName TooOld -NotePropertyValue $true; return $old }
    return $null
}

function Report($r, [int]$code, [string]$msg) {
    Write-Output ("PYTHON=" + ($(if ($r -and -not $r.TooOld) { $r.Path } else { "" })))
    Write-Output ("VERSION=" + ($(if ($r) { $r.Version } else { "" })))
    if ($msg) { Write-Output $msg }
    exit $code
}

$found = Find-Python
if ($found -and -not $found.TooOld) {
    Report $found 0 "[check_python] 사용 가능: Python $($found.Version) — $($found.Path)"
}

if (-not $Install) {
    $why = if ($found) { "Python $($found.Version) 만 있습니다(3.10 이상 필요 — $($found.Path))" } else { "Python 3.10 이상을 찾지 못했습니다" }
    $how = if (Get-Command winget) {
        "설치 방법: ① 동의하면 이 스크립트를 -Install 로 다시 실행(winget, 현재 사용자에게만 설치) " +
        "② 직접 설치 https://www.python.org/downloads/ — 설치 첫 화면에서 'Add python.exe to PATH' 체크"
    } else {
        "설치 방법: https://www.python.org/downloads/ 에서 직접 설치 — 설치 첫 화면에서 " +
        "'Add python.exe to PATH' 체크 (이 PC 에는 winget 이 없어 자동 설치를 할 수 없습니다)"
    }
    Report $found $(if ($found) { 2 } else { 1 }) "[check_python] $why`n[check_python] $how"
}

# --- 설치(사용자 동의 후) ---
if (-not (Get-Command winget)) {
    Report $null 3 "[check_python] winget 이 없어 자동 설치를 할 수 없습니다 — https://www.python.org/downloads/ 에서 직접 설치하세요"
}
Write-Output "[check_python] winget 으로 Python 3.12 를 현재 사용자에게 설치합니다(수 분 걸릴 수 있음)..."
winget install -e --id Python.Python.3.12 --scope user --silent `
    --accept-package-agreements --accept-source-agreements | Out-Null
$exe = "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"
$r = if (Test-Path $exe) { Probe $exe } else { Find-Python }
if ($r -and -not $r.TooOld) {
    Report $r 0 ("[check_python] 설치 완료: Python $($r.Version) — $($r.Path)`n" +
                 "[check_python] 이 창에는 PATH 가 아직 반영되지 않았습니다 — 위 경로로 이어서 실행하세요")
}
Report $null 3 "[check_python] 설치에 실패했습니다(winget 종료 코드 $LASTEXITCODE) — https://www.python.org/downloads/ 에서 직접 설치하세요"
