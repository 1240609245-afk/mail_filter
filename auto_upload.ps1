$ErrorActionPreference = "Stop"

$workDir = "C:\Users\li\Desktop\mail_filter"
$pythonScript = "mail_filter.py"
$logFile = Join-Path $workDir "task_log.txt"

Start-Transcript -Path $logFile -Append

try {
    Set-Location $workDir
    Write-Host "Current folder: $workDir"

    $pythonCmd = $null
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $pythonCmd = "py"
    }
    elseif (Get-Command python -ErrorAction SilentlyContinue) {
        $pythonCmd = "python"
    }
    else {
        throw "Python not found. Please install Python and add to PATH."
    }

    Write-Host "Using Python: $pythonCmd"
    Write-Host "Running script: $pythonScript"

    if ($pythonCmd -eq "py") {
        & py $pythonScript
    } else {
        & python $pythonScript
    }

    if ($LASTEXITCODE -ne 0) {
        throw "Python script failed, code: $LASTEXITCODE"
    }

    Write-Host "Python done"

    if (Get-Command git -ErrorAction SilentlyContinue) {
        Write-Host "Running git upload"

        & git add .
        & git commit -m "auto update"
        & git push

        Write-Host "Git done"
    }
    else {
        throw "Git not found"
    }

    Write-Host "All done"
}
catch {
    Write-Host "Error:"
    Write-Host $_
    exit 1
}
finally {
    Stop-Transcript
}