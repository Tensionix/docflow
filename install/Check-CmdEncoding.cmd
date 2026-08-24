@echo off
chcp 65001 >nul
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
for %%A in ("%SCRIPT_DIR%\..") do set "ROOT=%%~fA"

set "CHECKER=%ROOT%\system_core\core\cmd_encoding.py"
if not exist "%CHECKER%" goto ERR_CHECKER

call :RESOLVE_PYTHON
if errorlevel 1 goto ERR_PYTHON

set "CHECK_ARGS="

:ARGS
if "%~1"=="" goto RUN
if /I "%~1"=="-Fix" (
  set "CHECK_ARGS=%CHECK_ARGS% --fix"
) else if /I "%~1"=="/Fix" (
  set "CHECK_ARGS=%CHECK_ARGS% --fix"
) else if /I "%~1"=="--fix" (
  set "CHECK_ARGS=%CHECK_ARGS% --fix"
) else (
  set "CHECK_ARGS=%CHECK_ARGS% %~1"
)
shift
goto ARGS

:RUN
"%PYTHON_CMD%" %PYTHON_ARGS% "%CHECKER%" --root "%ROOT%" %CHECK_ARGS%
exit /b %errorlevel%

:RESOLVE_PYTHON
set "PYTHON_CMD="
set "PYTHON_ARGS="

if exist "%ROOT%\runtime\python.exe" (
  set "PYTHON_CMD=%ROOT%\runtime\python.exe"
  exit /b 0
)

if exist "%ROOT%\runtime\python\python.exe" (
  set "PYTHON_CMD=%ROOT%\runtime\python\python.exe"
  exit /b 0
)

py -3.12 -V >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=py"
  set "PYTHON_ARGS=-3.12"
  exit /b 0
)

python --version >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=python"
  exit /b 0
)

exit /b 1

:ERR_CHECKER
echo [ERROR] CMD encoding checker was not found:
echo %CHECKER%
exit /b 1

:ERR_PYTHON
echo [ERROR] Python was not found for CMD encoding check.
exit /b 1
