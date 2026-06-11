$ErrorActionPreference = 'Continue'
$repo       = 'C:\Users\eklementeva\Cowork\stm-vercel-dashboard'
$deployRepo = 'C:\Users\eklementeva\Cowork\hub-stm-deploy'
$deployRel  = 'hub-v2/public/stm-dashboard/index.html'
Set-Location $repo

$ts  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$log = Join-Path $repo 'daily_refresh.log'
"=== $ts START ==="                                       | Tee-Object -FilePath $log -Append | Out-Null

# подчищаем артефакты прошлого упавшего билда, чтобы git pull --ff-only не падал
git checkout -- index.html 2>&1                           | Tee-Object -FilePath $log -Append | Out-Null

git pull --ff-only 2>&1                                   | Tee-Object -FilePath $log -Append | Out-Null
if ($LASTEXITCODE -ne 0) {
    "=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') GIT PULL FAILED ===" | Tee-Object -FilePath $log -Append | Out-Null
    exit 1
}

& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'scripts\build_stm_data.ps1') 2>&1 |
    Tee-Object -FilePath $log -Append | Out-Null
$buildExit = $LASTEXITCODE
"build exit: $buildExit"                                  | Tee-Object -FilePath $log -Append | Out-Null

if ($buildExit -ne 0) {
    "=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') BUILD FAILED ==="  | Tee-Object -FilePath $log -Append | Out-Null
    exit 1
}

# 1) Архивный коммит в KaterZakharova/stm-vercel-dashboard (история данных + fallback Vercel зеркало)
git add index.html 2>&1 | Tee-Object -FilePath $log -Append | Out-Null
git diff --cached --quiet
$hasChange = ($LASTEXITCODE -ne 0)
if ($hasChange) {
    $today = Get-Date -Format 'dd.MM.yyyy'
    git commit -m "data: auto-refresh $today" 2>&1 | Tee-Object -FilePath $log -Append | Out-Null
    git push 2>&1                                  | Tee-Object -FilePath $log -Append | Out-Null
    "archive push exit: $LASTEXITCODE"             | Tee-Object -FilePath $log -Append | Out-Null
} else {
    "index.html не изменился — пропускаю архивный push" | Tee-Object -FilePath $log -Append | Out-Null
}

# 2) Боевой деплой: push в GSS-AI-Native/hub-stm → Coolify пересоберёт stm-v2.gsscosmetics.com/stm-dashboard
"--- deploy to hub-stm (Coolify) ---" | Tee-Object -FilePath $log -Append | Out-Null
Push-Location $deployRepo
try {
    # привести clone в чистое состояние origin/main (на всякий случай)
    git fetch origin main 2>&1                  | Tee-Object -FilePath $log -Append | Out-Null
    git checkout main 2>&1                      | Tee-Object -FilePath $log -Append | Out-Null
    git reset --hard origin/main 2>&1           | Tee-Object -FilePath $log -Append | Out-Null
    git clean -fd 2>&1                          | Tee-Object -FilePath $log -Append | Out-Null

    Copy-Item -Path (Join-Path $repo 'index.html') -Destination (Join-Path $deployRepo $deployRel) -Force
    git add $deployRel 2>&1                     | Tee-Object -FilePath $log -Append | Out-Null
    git diff --cached --quiet
    if ($LASTEXITCODE -ne 0) {
        $today = Get-Date -Format 'dd.MM.yyyy'
        git commit -m "hub-v2/stm-dashboard: auto-refresh $today" 2>&1 | Tee-Object -FilePath $log -Append | Out-Null
        git push origin main 2>&1                | Tee-Object -FilePath $log -Append | Out-Null
        "deploy push exit: $LASTEXITCODE"        | Tee-Object -FilePath $log -Append | Out-Null
    } else {
        "stm-dashboard/index.html в hub-stm не изменился — пропускаю прод-push" | Tee-Object -FilePath $log -Append | Out-Null
    }
}
finally {
    Pop-Location
}

"=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') DONE ===" | Tee-Object -FilePath $log -Append | Out-Null
