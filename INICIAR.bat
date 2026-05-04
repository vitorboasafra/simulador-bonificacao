@echo off
setlocal
cd /d "%~dp0"

echo =========================================
echo Iniciando o sistema Boa Safra em Node.js...
echo =========================================

where node >nul 2>&1
if errorlevel 1 (
  echo [ERRO] Node.js nao encontrado no PATH.
  pause
  exit /b 1
)

where npm.cmd >nul 2>&1
if errorlevel 1 (
  echo [ERRO] npm.cmd nao encontrado no PATH.
  pause
  exit /b 1
)

if not exist "package.json" (
  echo [ERRO] package.json nao encontrado nesta pasta.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo Instalando dependencias Node.js...
  call npm.cmd install
  if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias com npm.
    pause
    exit /b 1
  )
)

if not exist "node_modules\multer" (
  echo Atualizando dependencias Node.js...
  call npm.cmd install
  if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias com npm.
    pause
    exit /b 1
  )
)

if not exist "node_modules\pdf-lib" (
  echo Instalando dependencias Node.js...
  call npm.cmd install
  if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias com npm.
    pause
    exit /b 1
  )
)

for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":3000" ^| findstr "LISTENING"') do (
  echo Encerrando processo anterior na porta 3000 (PID %%P)...
  taskkill /PID %%P /F >nul 2>&1
)

start "Boa Safra - Login Node" cmd /k "cd /d ""%~dp0"" && npm.cmd run dev"

timeout /t 3 /nobreak >nul
start "" "http://localhost:3000"

echo Sistema iniciado. Para encerrar, feche a janela "Boa Safra - Login Node".
pause
