@echo off
chcp 65001 >nul
echo ===================================================================
echo TU DONG BIEN DICH VA DAY TRUC TIEP LEN GITHUB (anhtrung4566/software-hub)
echo ===================================================================
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:winexe /out:"%~dp0downloads\VBTM-Setup-2.2.4.exe" "%~dp0desktop-src\VBTM-Pro\VBTMSetupAndApp.cs"
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:winexe /out:"%~dp0downloads\VietNamOCR-Setup-3.2.0-x64.exe" "%~dp0desktop-src\VietNamOCR-Pro\VietNamOCRSetupAndApp.cs"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0push_to_github.ps1"
pause
