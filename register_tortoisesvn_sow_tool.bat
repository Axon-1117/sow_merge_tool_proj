@echo off
setlocal enabledelayedexpansion

set "TOOL=D:\Tools\sow_merge_tool_proj\dist\sow_merge_tool.exe"
set "ROOT=HKCU\Software\TortoiseSVN"
set "DIFF_CMD=\"%TOOL%\" %%base %%mine --title %%bname"
set "MERGE_CMD=\"%TOOL%\" --base %%base --mine %%mine --theirs %%theirs --merged %%merged --title %%bname"
set "MERGE_ARGS=--base %%base --mine %%mine --theirs %%theirs --merged %%merged --title %%bname"

echo [1/4] Checking tool path...
if not exist "%TOOL%" (
  echo ERROR: tool not found: %TOOL%
  exit /b 1
)

echo [2/4] Setting global Diff/Merge...
reg add "%ROOT%" /v Diff /t REG_SZ /d "%DIFF_CMD%" /f >nul
reg add "%ROOT%" /v Merge /t REG_SZ /d "%MERGE_CMD%" /f >nul
reg add "%ROOT%" /v DiffArgs /t REG_SZ /d "" /f >nul
reg add "%ROOT%" /v MergeArgs /t REG_SZ /d "" /f >nul

echo [3/4] Setting per-extension tools (.xlsx/.xlsm/.csv)...
for %%E in (.xlsx .xlsm .csv) do (
  reg add "%ROOT%\DiffTools" /v %%E /t REG_SZ /d "%DIFF_CMD%" /f >nul
  reg add "%ROOT%\MergeTools" /v %%E /t REG_SZ /d "%MERGE_CMD%" /f >nul
)

echo [4/4] Setting command/args under extension subkeys...
for %%K in (XLSX XLSM CSV) do (
  reg add "%ROOT%\DiffTools\%%K" /v command /t REG_SZ /d "%DIFF_CMD%" /f >nul
  reg add "%ROOT%\DiffTools\%%K" /v args /t REG_SZ /d "%%base %%mine --title %%bname" /f >nul
  reg add "%ROOT%\MergeTools\%%K" /v command /t REG_SZ /d "%MERGE_CMD%" /f >nul
  reg add "%ROOT%\MergeTools\%%K" /v args /t REG_SZ /d "%MERGE_ARGS%" /f >nul
)

echo Done.
echo Diff/Merge tool has been updated to:
echo   %TOOL%
exit /b 0

