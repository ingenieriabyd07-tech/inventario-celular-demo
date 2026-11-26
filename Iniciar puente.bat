@echo off
chcp 65001 >nul
title Puente Asepsia Total
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (python puente\servidor.py %*) else (py puente\servidor.py %*)
pause
