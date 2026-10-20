$ErrorActionPreference = 'Stop'
$proj = Split-Path -Parent $MyInvocation.MyCommand.Definition

function Get-Pyw {
    try {
        $p = & py -3 -c "import sys,os;print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))" 2>$null
        if ($p -and (Test-Path $p)) { return $p }
    } catch {}
    foreach ($c in @('pythonw','pyw')) {
        try {
            $which = (Get-Command $c -ErrorAction Stop).Source
            if ($which) { return $which }
        } catch {}
    }
    return $null
}

$pyw = Get-Pyw
if (-not $pyw) {
    Write-Host "Cannot find pythonw. Please run run.bat once to install dependencies first."
    exit 1
}

# v128: point the shortcut at watchdog.py (the supervisor), not main.py.
# The supervisor keeps the main program alive and restarts it after a crash or
# a COM hang -- otherwise a hang leaves the single-instance mutex held and the
# next launch silently does nothing until the machine is rebooted.
$target = Join-Path $proj 'watchdog.py'
$startup = [System.Environment]::GetFolderPath('Startup')
$lnkPath = Join-Path $startup 'VbeComplete.lnk'

$shell = New-Object -ComObject WScript.Shell
$lnk = $shell.CreateShortcut($lnkPath)
$lnk.TargetPath = $pyw
$lnk.Arguments = "`"$target`""
$lnk.WorkingDirectory = $proj
$lnk.WindowStyle = 7
$lnk.Description = 'VbeComplete - VBE auto-complete (with watchdog)'
$lnk.Save()

Write-Host ("Created startup shortcut: " + $lnkPath)
Write-Host "VbeComplete will start automatically when you log in."
