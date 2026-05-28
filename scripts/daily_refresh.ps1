$ErrorActionPreference = 'Continue'
$repo = 'C:\Users\eklementeva\Cowork\stm-vercel-dashboard'
Set-Location $repo

$ts  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$log = Join-Path $repo 'daily_refresh.log'
"=== $ts START ==="                                       | Tee-Object -FilePath $log -Append | Out-Null

git pull --ff-only 2>&1                                   | Tee-Object -FilePath $log -Append | Out-Null

& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'scripts\build_stm_data.ps1') 2>&1 |
    Tee-Object -FilePath $log -Append | Out-Null
$buildExit = $LASTEXITCODE
"build exit: $buildExit"                                  | Tee-Object -FilePath $log -Append | Out-Null

if ($buildExit -ne 0) {
    "=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') BUILD FAILED ==="  | Tee-Object -FilePath $log -Append | Out-Null
    exit 1
}

git add index.html 2>&1 | Tee-Object -FilePath $log -Append | Out-Null
git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
    $today = Get-Date -Format 'dd.MM.yyyy'
    git commit -m "data: auto-refresh $today" 2>&1 | Tee-Object -FilePath $log -Append | Out-Null
    git push 2>&1                                  | Tee-Object -FilePath $log -Append | Out-Null
    & "$env:APPDATA\npm\vercel.cmd" --prod --yes 2>&1 | Tee-Object -FilePath $log -Append | Out-Null
    "deploy exit: $LASTEXITCODE"                    | Tee-Object -FilePath $log -Append | Out-Null
} else {
    "index.html не изменился — пропускаю commit/deploy" | Tee-Object -FilePath $log -Append | Out-Null
}

"=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') DONE ===" | Tee-Object -FilePath $log -Append | Out-Null
