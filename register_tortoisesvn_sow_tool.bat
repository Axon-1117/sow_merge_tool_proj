@echo off
setlocal enabledelayedexpansion

set "TOOL=D:\Tools\sow_merge_tool_proj\dist\sow_merge_tool.exe"
set "ROOT=HKCU\Software\TortoiseSVN"
set "DIFFTOOLS=%ROOT%\DiffTools"
set "MERGETOOLS=%ROOT%\MergeTools"
set "DIFF_CMD=\"%TOOL%\" %%base %%mine --title %%bname"
set "MERGE_CMD=\"%TOOL%\" --base %%base --mine %%mine --theirs %%theirs --merged %%merged --title %%bname"
set "MERGE_ARGS=--base %%base --mine %%mine --theirs %%theirs --merged %%merged --title %%bname"

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TS=%%I"
set "BACKUP_ROOT=%~dp0backups"
set "BACKUP_DIR=%BACKUP_ROOT%\%TS%"
set "REG_MAIN=%BACKUP_DIR%\TortoiseSVN.reg"
set "REG_DIFF=%BACKUP_DIR%\TortoiseSVN_DiffTools.reg"
set "REG_MERGE=%BACKUP_DIR%\TortoiseSVN_MergeTools.reg"
set "RESTORE_BAT=%BACKUP_DIR%\restore_tortoisesvn_config_%TS%.bat"
set "LATEST_RESTORE=%~dp0restore_tortoisesvn_config_latest.bat"

echo [1/5] Checking tool path...
if not exist "%TOOL%" (
  echo ERROR: tool not found: %TOOL%
  exit /b 1
)

echo [2/5] Backing up current TortoiseSVN Diff/Merge settings...
if not exist "%BACKUP_ROOT%" mkdir "%BACKUP_ROOT%" >nul 2>nul
if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%" >nul 2>nul

reg export "%ROOT%" "%REG_MAIN%" /y >nul 2>nul
reg export "%DIFFTOOLS%" "%REG_DIFF%" /y >nul 2>nul
reg export "%MERGETOOLS%" "%REG_MERGE%" /y >nul 2>nul

(
  echo @echo off
  echo setlocal
  echo echo Restoring TortoiseSVN Diff/Merge settings from:
  echo echo   %BACKUP_DIR%
  echo if exist "%REG_MAIN%" reg import "%REG_MAIN%"
  echo if exist "%REG_DIFF%" reg import "%REG_DIFF%"
  echo if exist "%REG_MERGE%" reg import "%REG_MERGE%"
  echo echo Done.
  echo pause
  echo exit /b 0
) > "%RESTORE_BAT%"
copy /y "%RESTORE_BAT%" "%LATEST_RESTORE%" >nul

echo [3/5] Setting global Diff/Merge...
reg add "%ROOT%" /v Diff /t REG_SZ /d "%DIFF_CMD%" /f >nul
reg add "%ROOT%" /v Merge /t REG_SZ /d "%MERGE_CMD%" /f >nul
reg add "%ROOT%" /v DiffArgs /t REG_SZ /d "" /f >nul
reg add "%ROOT%" /v MergeArgs /t REG_SZ /d "" /f >nul

echo [4/5] Setting per-extension tools (.xlsx/.xlsm/.csv)...
for %%E in (.xlsx .xlsm .csv) do (
  reg add "%DIFFTOOLS%" /v %%E /t REG_SZ /d "%DIFF_CMD%" /f >nul
  reg add "%MERGETOOLS%" /v %%E /t REG_SZ /d "%MERGE_CMD%" /f >nul
)

echo [5/5] Setting command/args under extension subkeys...
for %%K in (XLSX XLSM CSV) do (
  reg add "%DIFFTOOLS%\%%K" /v command /t REG_SZ /d "%DIFF_CMD%" /f >nul
  reg add "%DIFFTOOLS%\%%K" /v args /t REG_SZ /d "%%base %%mine --title %%bname" /f >nul
  reg add "%MERGETOOLS%\%%K" /v command /t REG_SZ /d "%MERGE_CMD%" /f >nul
  reg add "%MERGETOOLS%\%%K" /v args /t REG_SZ /d "%MERGE_ARGS%" /f >nul
)

echo Done.
echo Diff/Merge tool has been updated to:
echo   %TOOL%
echo Backup saved at:
echo   %BACKUP_DIR%
echo Restore script:
echo   %RESTORE_BAT%
echo Latest restore shortcut:
echo   %LATEST_RESTORE%
exit /b 0
