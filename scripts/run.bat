@echo off
REM coffice2txt - Windows launcher
REM
REM Usage:
REM   run.bat [file.doc^|file.xls^|file.ppt^|file.docx^|file.xlsx^|file.pptx]

setlocal
cd /d "%~dp0"

where perl >nul 2>&1
if errorlevel 1 (
    echo perl was not found on PATH.
    echo Install Strawberry Perl from https://strawberryperl.com/
    pause
    exit /b 1
)

set "PERL5LIB=%CD%;%CD%\Spreadsheet-XLSX-TempFolderCreator\lib;%CD%\Spreadsheet-XLSX-Utility2007PDP\lib;%CD%\Spreadsheet-XLSX-ZipMod2\lib;%PERL5LIB%"

perl coffice2txt.pl %*
endlocal
