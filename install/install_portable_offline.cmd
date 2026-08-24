@echo off
setlocal EnableExtensions

title Audion DocFlow - Offline Install

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
for %%A in ("%SCRIPT_DIR%\..") do set "ROOT=%%~fA"

set "RUNTIME_DIR=%ROOT%\runtime"
set "WHEELHOUSE_DIR=%ROOT%\wheelhouse"
set "REQ_FILE=%ROOT%\install\requirements_full.in"
set "GET_PIP=%ROOT%\install\download\get-pip.py"

echo ======================================================================
echo   AUDION DOCFLOW - OFFLINE INSTALL
echo ======================================================================
echo Root:       %ROOT%
echo Runtime:    %RUNTIME_DIR%
echo Wheelhouse: %WHEELHOUSE_DIR%
echo.

if not exist "%RUNTIME_DIR%\python.exe" goto ERR_PYTHON
if not exist "%REQ_FILE%" goto ERR_REQ

"%RUNTIME_DIR%\python.exe" -m pip --version >nul 2>nul
if errorlevel 1 goto NEED_PIP
goto INSTALL

:NEED_PIP
if not exist "%GET_PIP%" goto ERR_GETPIP
echo [1/3] Installing pip from cached get-pip.py...
"%RUNTIME_DIR%\python.exe" "%GET_PIP%"
if errorlevel 1 goto ERR_PIP

:INSTALL
echo [2/3] Installing build bootstrap from wheelhouse...
"%RUNTIME_DIR%\python.exe" -m pip install --disable-pip-version-check --no-index --find-links="%WHEELHOUSE_DIR%" setuptools wheel packaging
if errorlevel 1 goto ERR_INSTALL

echo [2/3] Installing packages from wheelhouse...
"%RUNTIME_DIR%\python.exe" -m pip install --disable-pip-version-check --no-build-isolation --no-index --find-links="%WHEELHOUSE_DIR%" -r "%REQ_FILE%"
if errorlevel 1 goto ERR_INSTALL

echo [3/3] Verifying environment...
call "%ROOT%\install\verify_portable_env.cmd"
exit /b %errorlevel%

:ERR_PYTHON
echo [ERROR] runtime\python.exe was not found.
if not defined AUDION_NO_PAUSE pause
exit /b 1

:ERR_REQ
echo [ERROR] requirements_full.in was not found.
if not defined AUDION_NO_PAUSE pause
exit /b 1

:ERR_GETPIP
echo [ERROR] Cached get-pip.py was not found.
echo [INFO] Run Build_Portable_Env_Build.cmd once on an online machine.
if not defined AUDION_NO_PAUSE pause
exit /b 1

:ERR_PIP
echo [ERROR] Failed to install pip.
if not defined AUDION_NO_PAUSE pause
exit /b 1

:ERR_INSTALL
echo [ERROR] Offline install failed.
if not defined AUDION_NO_PAUSE pause
exit /b 1
