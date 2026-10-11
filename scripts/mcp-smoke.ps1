# MCP 冒烟测试：向 mcp_server.py 发送 initialize + tools/list，检查返回
$ErrorActionPreference = 'Stop'
$server = Join-Path $PSScriptRoot '..\tools\vba\mcp_server.py'
$out    = Join-Path $env:TEMP 'mcp-smoke.out'

$reqs = @(
  '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"smoke","version":"1"}}}',
  '{"jsonrpc":"2.0","method":"notifications/initialized"}',
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}'
)

$reqs | python $server 2>&1 | Out-File -Encoding utf8 $out

Write-Host "exit=$LASTEXITCODE  lines=$((Get-Content $out).Count)"
Get-Content $out | ForEach-Object {
    try { $_ | ConvertFrom-Json | Out-Null; $ok = 'JSON-OK' } catch { $ok = 'NOT-JSON' }
    "{0}  {1}" -f $ok, $_.Substring(0, [Math]::Min(160, $_.Length))
}
