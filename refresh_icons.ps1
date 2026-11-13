param(
    [Parameter(Mandatory = $true)][string]$IconPath,
    [Parameter(Mandatory = $true)][string]$ExePath
)

$shell = New-Object -ComObject WScript.Shell
$iconLocation = "$IconPath,0"
$workDir = Split-Path -Parent $ExePath
$updated = New-Object System.Collections.Generic.List[string]

function Save-AppShortcut([string]$Path, [bool]$RequireOurTarget) {
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $lnk = $shell.CreateShortcut($Path)
    $target = [string]$lnk.TargetPath
    $named = [IO.Path]::GetFileName($Path) -eq "Debt Payoff Planner.lnk"
    $ours = [string]::Equals($target, $ExePath, [StringComparison]::OrdinalIgnoreCase)
    if ($RequireOurTarget -and -not $ours) { return }
    if (-not $named -and -not $ours) { return }
    try {
        $lnk.TargetPath = $ExePath
        $lnk.WorkingDirectory = $workDir
        $lnk.IconLocation = $iconLocation
        $lnk.Save()
        if (-not $updated.Contains($Path)) { $updated.Add($Path) }
    } catch {}
}

$dirs = @(
    "$env:PUBLIC\Desktop",
    [Environment]::GetFolderPath("Desktop"),
    "$env:USERPROFILE\Desktop",
    "$env:USERPROFILE\OneDrive\Desktop",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Debt Payoff Planner",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Debt Payoff Planner",
    "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar",
    "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu"
)

foreach ($dir in $dirs) {
    if (-not (Test-Path -LiteralPath $dir)) { continue }
    $pinned = $dir -like "*User Pinned*"
    Get-ChildItem -LiteralPath $dir -Filter *.lnk -ErrorAction SilentlyContinue | ForEach-Object {
        Save-AppShortcut $_.FullName $pinned
    }
}

$required = @(
    "$env:PUBLIC\Desktop\Debt Payoff Planner.lnk",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Debt Payoff Planner\Debt Payoff Planner.lnk"
)
foreach ($path in $required) {
    $parent = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    try {
        $lnk = $shell.CreateShortcut($path)
        $lnk.TargetPath = $ExePath
        $lnk.WorkingDirectory = $workDir
        $lnk.IconLocation = $iconLocation
        $lnk.Description = "Debt Payoff Planner"
        $lnk.Save()
        if (-not $updated.Contains($path)) { $updated.Add($path) }
    } catch {}
}

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class DebtIconNote {
  [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
  public static extern void SHChangeNotify(uint ev, uint flags, string item, IntPtr item2);
  [DllImport("shell32.dll")]
  public static extern void SHChangeNotify(uint ev, uint flags, IntPtr item, IntPtr item2);
}
"@

foreach ($path in $updated) {
    [DebtIconNote]::SHChangeNotify(0x2000, 0x1005, $path, [IntPtr]::Zero)
}
[DebtIconNote]::SHChangeNotify(0x2000, 0x1005, $ExePath, [IntPtr]::Zero)
[DebtIconNote]::SHChangeNotify(0x2000, 0x1005, $IconPath, [IntPtr]::Zero)
[DebtIconNote]::SHChangeNotify(0x1000, 0x1005, "$env:PUBLIC\Desktop", [IntPtr]::Zero)
[DebtIconNote]::SHChangeNotify(0x08000000, 0x1000, [IntPtr]::Zero, [IntPtr]::Zero)

$ie4 = Join-Path $env:WINDIR "System32\ie4uinit.exe"
if (Test-Path -LiteralPath $ie4) {
    Start-Process -FilePath $ie4 -ArgumentList "-show" -WindowStyle Hidden -Wait
}
