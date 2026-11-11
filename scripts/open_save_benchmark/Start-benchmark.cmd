@echo off
setlocal
set "PSModulePath=%SystemRoot%\System32\WindowsPowerShell\v1.0\Modules"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\launch.ps1" -StudyRoot "%~dp0." %*
set "benchmark_exit=%errorlevel%"
pause
exit /b %benchmark_exit%
