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
        throw "Python not found. Please install Python and add it to PATH."
    }

    Write-Host "Using Python: $pythonCmd"
    Write-Host "Running script: $pythonScript"

    if ($pythonCmd -eq "py") {
        & py $pythonScript
    }
    else {
        & python $pythonScript
    }

    if ($LASTEXITCODE -ne 0) {
        throw "Python script failed, exit code: $LASTEXITCODE"
    }

    Write-Host "Python done"
    Write-Host "Running git upload"

    & git config --global http.version HTTP/1.1
    & git config --global http.postBuffer 524288000

    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw "Git not found. Please install Git and add it to PATH."
    }

    & git add .

    & git diff --cached --quiet
    if ($LASTEXITCODE -eq 0) {
        Write-Host "No changes to commit"
    }
    else {
        & git commit -m "auto update"
        if ($LASTEXITCODE -ne 0) {
            throw "git commit failed, exit code: $LASTEXITCODE"
        }
    }

    & git push origin main
    if ($LASTEXITCODE -ne 0) {
        throw "git push failed, exit code: $LASTEXITCODE"
    }

    Write-Host "Git done"
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