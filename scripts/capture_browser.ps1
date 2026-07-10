# 화면에 보이는 브라우저 창 영역을 PNG로 캡처한다 (의존성 없는 폴백 수단).
# 사전조건: 대상 창이 화면에 보이는 상태(최소화 아님)여야 한다.
#
# 사용 예:
#   powershell -ExecutionPolicy Bypass -File capture_browser.ps1 -OutFile shot.png
#   powershell -ExecutionPolicy Bypass -File capture_browser.ps1 -OutFile shot.png -WindowTitlePattern "Edge"

param(
    [Parameter(Mandatory = $true)][string]$OutFile,
    [string]$WindowTitlePattern = "Chrome",
    [int]$DelayMs = 500
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32Capture {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr hwnd, int attr, out RECT rect, int size);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
}
"@

# DPI 배율(125%/150%) 환경에서 캡처 영역이 어긋나는 것을 막는다 - 반드시 좌표 취득 전에 호출
[void][Win32Capture]::SetProcessDPIAware()

# 제목 패턴으로 대상 창 탐색
$proc = Get-Process |
    Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero -and $_.MainWindowTitle -match $WindowTitlePattern } |
    Select-Object -First 1
if (-not $proc) {
    Write-Error "window not found: title pattern '$WindowTitlePattern'. Open the target browser window first."
    exit 1
}
$hwnd = $proc.MainWindowHandle

# 최소화 상태면 복원(SW_RESTORE=9) 후 전면으로 가져와 캡처 준비
if ([Win32Capture]::IsIconic($hwnd)) { [void][Win32Capture]::ShowWindow($hwnd, 9) }
[void][Win32Capture]::SetForegroundWindow($hwnd)
Start-Sleep -Milliseconds $DelayMs

# 그림자 제외 실제 창 영역(DWMWA_EXTENDED_FRAME_BOUNDS=9); 실패 시 GetWindowRect 폴백
$rect = New-Object "Win32Capture+RECT"
$size = [System.Runtime.InteropServices.Marshal]::SizeOf([type]"Win32Capture+RECT")
$hr = [Win32Capture]::DwmGetWindowAttribute($hwnd, 9, [ref]$rect, $size)
if ($hr -ne 0) { [void][Win32Capture]::GetWindowRect($hwnd, [ref]$rect) }

$w = $rect.Right - $rect.Left
$h = $rect.Bottom - $rect.Top
if ($w -le 0 -or $h -le 0) {
    Write-Error "invalid window rect ($w x $h). The window may be minimized or off-screen."
    exit 1
}

$dir = Split-Path -Parent $OutFile
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

$bmp = New-Object System.Drawing.Bitmap($w, $h)
$g = [System.Drawing.Graphics]::FromImage($bmp)
try {
    $g.CopyFromScreen($rect.Left, $rect.Top, 0, 0, (New-Object System.Drawing.Size($w, $h)))
    $bmp.Save($OutFile, [System.Drawing.Imaging.ImageFormat]::Png)
}
finally {
    $g.Dispose()
    $bmp.Dispose()
}

Write-Output "[capture_browser] saved: $OutFile (${w}x${h}, window: $($proc.MainWindowTitle))"
