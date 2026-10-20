# Stop VbeComplete. Only python processes whose command line mentions
# VbeComplete are touched -- other Python programs are left alone.
#
# v128: the watchdog (watchdog.py) restarts the main program within a few
# seconds, so it MUST be killed first -- otherwise "stop" looks like it did
# nothing. We also run two rounds: killing the main program first can make the
# watchdog spawn a fresh one before it is itself killed.
$names = "Name = 'pythonw.exe' OR Name = 'python.exe' OR Name = 'py.exe' OR Name = 'pyw.exe'"
$n = 0

function Get-VbeProcs {
    Get-CimInstance Win32_Process -Filter $names -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -and $_.CommandLine -like "*VbeComplete*" }
}

for ($round = 1; $round -le 2; $round++) {
    $procs = @(Get-VbeProcs)
    if ($procs.Count -eq 0) { break }
    # Watchdog first: sort so that watchdog.py comes before everything else.
    $ordered = @($procs | Sort-Object -Property @{ Expression = { if ($_.CommandLine -like "*watchdog.py*") { 0 } else { 1 } } })
    foreach ($p in $ordered) {
        try {
            Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
            Write-Host ("killed PID " + $p.ProcessId)
            $n++
        } catch {
            Write-Host ("failed to kill PID " + $p.ProcessId + ": " + $_.Exception.Message)
        }
    }
    Start-Sleep -Milliseconds 600
}

if ($n -eq 0) {
    Write-Host "No VbeComplete process is running."
} else {
    Write-Host ("Stopped " + $n + " process(es).")
}
