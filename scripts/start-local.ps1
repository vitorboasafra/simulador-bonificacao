$ErrorActionPreference = "Stop"

$port = 3000
if ($env:PORT) {
  $parsedPort = 0
  if ([int]::TryParse($env:PORT, [ref]$parsedPort)) {
    $port = $parsedPort
  }
}

$connections = @(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue)
$processIds = @($connections | Select-Object -ExpandProperty OwningProcess -Unique)

foreach ($processId in $processIds) {
  if ($processId -and $processId -ne $PID) {
    Write-Host "Encerrando processo anterior na porta $port (PID $processId)..."
    Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
  }
}

if ($processIds.Count -gt 0) {
  Start-Sleep -Seconds 1
}

node server.js
