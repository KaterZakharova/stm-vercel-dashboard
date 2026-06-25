$ErrorActionPreference = 'Continue'
$repo     = 'C:\Users\eklementeva\Cowork\stm-vercel-dashboard'
$lockFile = Join-Path $repo 'refresh.lock'
$daily    = Join-Path $repo 'scripts\daily_refresh.ps1'

try {
    & $daily
} catch {
    "ERROR $($_.Exception.Message)" | Out-File -FilePath (Join-Path $repo 'refresh_runner.err.log') -Append -Encoding UTF8
} finally {
    Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
}
