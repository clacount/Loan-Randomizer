@echo off
setlocal
set "APP_HTML=%~dp0..\src\app\index.html"

if not exist "%APP_HTML%" (
  echo Could not find LendingFair at:
  echo %APP_HTML%
  pause
  exit /b 1
)

start "" "%APP_HTML%"
