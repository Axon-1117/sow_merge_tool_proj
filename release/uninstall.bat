@echo off
setlocal

echo Removing TortoiseSVN diff/merge tools for .xlsx ...

reg delete "HKCU\Software\TortoiseSVN\DiffTools" /v .xlsx /f
reg delete "HKCU\Software\TortoiseSVN\MergeTools" /v .xlsx /f

reg delete "HKCU\Software\TortoiseSVN\DiffTools\XLSX" /f
reg delete "HKCU\Software\TortoiseSVN\MergeTools\XLSX" /f

echo Done. Please restart TortoiseSVN / Explorer if needed.
pause
