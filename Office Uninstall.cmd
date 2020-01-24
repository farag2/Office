reg add "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f
cscript.exe "%~dp0Office Uninstall.vbs" ALL /Quiet /NoCancel /Force /OSE
reg add "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 0 /f
pause