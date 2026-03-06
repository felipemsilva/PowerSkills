<#
.SYNOPSIS
    System skill — shell commands, process info, environment.
#>
param(
    [string]$Action,
    [hashtable]$Args_,
    [hashtable]$Config
)

$defaultTimeout = if ($Config.default_timeout) { [int]$Config.default_timeout } else { 30 }

switch ($Action) {
    "exec" {
        $command = $Args_.command
        if (-not $command) { throw "Required: --command" }
        $timeout = if ($Args_.timeout) { [int]$Args_.timeout } else { $defaultTimeout }

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "powershell.exe"
        $psi.Arguments = "-NoProfile -NonInteractive -Command `"$($command -replace '"','\"')`""
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true

        $proc = [System.Diagnostics.Process]::Start($psi)
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()

        if (-not $proc.WaitForExit($timeout * 1000)) {
            $proc.Kill()
            return @{ stdout = ""; stderr = "TIMEOUT after ${timeout}s"; exit_code = 124 }
        }
        return @{ stdout = $stdoutTask.Result; stderr = $stderrTask.Result; exit_code = $proc.ExitCode }
    }
    "info" {
        return @{
            hostname = $env:COMPUTERNAME
            user     = $env:USERNAME
            domain   = $env:USERDOMAIN
            os       = (Get-CimInstance Win32_OperatingSystem).Caption
            arch     = $env:PROCESSOR_ARCHITECTURE
            ps_version = $PSVersionTable.PSVersion.ToString()
            uptime_hours = [math]::Round((Get-CimInstance Win32_OperatingSystem).LastBootUpTime.Subtract((Get-Date)).TotalHours * -1, 1)
        }
    }
    "processes" {
        $limit = if ($Args_.limit) { [int]$Args_.limit } else { 20 }
        $procs = Get-Process | Sort-Object CPU -Descending | Select-Object -First $limit | ForEach-Object {
            @{
                name = $_.ProcessName
                pid  = $_.Id
                cpu  = [math]::Round($_.CPU, 2)
                mem_mb = [math]::Round($_.WorkingSet64 / 1MB, 1)
            }
        }
        return $procs
    }
    "env" {
        $name = $Args_.name
        if (-not $name) { throw "Required: --name" }
        $val = [Environment]::GetEnvironmentVariable($name)
        return @{ name = $name; value = $val }
    }
    default {
        throw "Unknown action: $Action. Use: exec, info, processes, env"
    }
}
