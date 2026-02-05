@echo off
setlocal

set TOOL_PATH=%~dp0sow_merge_tool.exe

echo Installing TortoiseSVN diff/merge tools for .xlsx ...

reg add "HKCU\Software\TortoiseSVN\DiffTools" /v .xlsx /t REG_SZ /d "\"%TOOL_PATH%\" --base \"%base\" --mine \"%mine\" --title \"%bname\"" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools" /v .xlsx /t REG_SZ /d "\"%TOOL_PATH%\" --base \"%base\" --mine \"%mine\" --theirs \"%theirs\" --merged \"%merged\" --title \"%bname\"" /f

reg add "HKCU\Software\TortoiseSVN\DiffTools\XLSX" /v command /t REG_SZ /d "%TOOL_PATH%" /f
reg add "HKCU\Software\TortoiseSVN\DiffTools\XLSX" /v args /t REG_SZ /d "--base %base --mine %mine --title %bname" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools\XLSX" /v command /t REG_SZ /d "%TOOL_PATH%" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools\XLSX" /v args /t REG_SZ /d "--base %base --mine %mine --theirs %theirs --merged %merged --title %bname" /f

echo Done. Please restart TortoiseSVN / Explorer if needed.
pause
