@echo off
chcp 65001 >nul
title Sieu App Van Phong - Localhost Server (Port 8080)
powershell -ExecutionPolicy Bypass -File "%~dp0Start-LocalServer.ps1" -Port 8080
pause
