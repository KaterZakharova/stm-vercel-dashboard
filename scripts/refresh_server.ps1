$ErrorActionPreference = 'Continue'
$repo     = 'C:\Users\eklementeva\Cowork\stm-vercel-dashboard'
$lockFile = Join-Path $repo 'refresh.lock'
$logFile  = Join-Path $repo 'daily_refresh.log'
$runner   = Join-Path $repo 'scripts\refresh_runner.ps1'
$srvLog   = Join-Path $repo 'refresh_server.log'

function Write-SrvLog($msg) {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts $msg" | Out-File -FilePath $srvLog -Append -Encoding UTF8
}

function Get-StatusJson {
    $running = Test-Path $lockFile
    $lastRunIso = ''
    $lastRunHuman = ''
    if (Test-Path $logFile) {
        $t = (Get-Item $logFile).LastWriteTime
        $lastRunIso = $t.ToString('yyyy-MM-ddTHH:mm:ss')
        $lastRunHuman = $t.ToString('dd.MM.yyyy HH:mm')
    }
    $startedAtIso = ''
    if ($running) {
        try { $startedAtIso = (Get-Content $lockFile -Raw -ErrorAction SilentlyContinue).Trim() } catch { }
    }
    return (@{
        running       = [bool]$running
        lastRun       = $lastRunIso
        lastRunHuman  = $lastRunHuman
        startedAt     = $startedAtIso
    } | ConvertTo-Json -Compress)
}

function Start-Refresh {
    if (Test-Path $lockFile) { return '{"status":"already_running"}' }
    Set-Content -Path $lockFile -Value (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss') -Encoding UTF8
    Start-Process -FilePath 'powershell.exe' `
        -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-WindowStyle','Hidden','-File',$runner `
        -WindowStyle Hidden
    Write-SrvLog 'refresh started'
    return '{"status":"started"}'
}

$http = [System.Net.HttpListener]::new()
$http.Prefixes.Add('http://127.0.0.1:9999/')
try {
    $http.Start()
    Write-SrvLog 'listening on http://127.0.0.1:9999/'
} catch {
    Write-SrvLog "start failed: $_"
    exit 1
}

while ($http.IsListening) {
    try {
        $ctx = $http.GetContext()
    } catch {
        Write-SrvLog "GetContext error: $_"
        break
    }
    $req = $ctx.Request
    $res = $ctx.Response
    try {
        $res.Headers.Add('Access-Control-Allow-Origin', '*')
        $res.Headers.Add('Cache-Control', 'no-store')
        $body = ''
        $status = 200

        if ($req.HttpMethod -eq 'OPTIONS') {
            $res.Headers.Add('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
            $res.Headers.Add('Access-Control-Allow-Headers', 'Content-Type')
            $res.Headers.Add('Access-Control-Max-Age', '86400')
            $status = 204
        }
        elseif ($req.Url.LocalPath -eq '/refresh' -and $req.HttpMethod -eq 'POST') {
            $body = Start-Refresh
        }
        elseif ($req.Url.LocalPath -eq '/refresh' -and $req.HttpMethod -eq 'GET') {
            # GET вариант — на случай дёрнуть из браузера руками
            $body = Start-Refresh
        }
        elseif ($req.Url.LocalPath -eq '/status') {
            $body = Get-StatusJson
        }
        elseif ($req.Url.LocalPath -eq '/' -or $req.Url.LocalPath -eq '/health') {
            $body = '{"ok":true,"service":"stm-refresh-server"}'
        }
        else {
            $status = 404
            $body = '{"error":"not_found"}'
        }

        $res.StatusCode = $status
        if ($body) {
            $buf = [System.Text.Encoding]::UTF8.GetBytes($body)
            $res.ContentType = 'application/json; charset=utf-8'
            $res.ContentLength64 = $buf.Length
            $res.OutputStream.Write($buf, 0, $buf.Length)
        }
    } catch {
        Write-SrvLog "handler error: $_"
        try { $res.StatusCode = 500 } catch {}
    } finally {
        try { $res.Close() } catch {}
    }
}
