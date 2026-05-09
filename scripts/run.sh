#!/usr/bin/env bash
# coffice2txt — Linux / ChromeOS launcher
#
# Usage:
#   ./run.sh [file.doc|file.xls|file.ppt|file.docx|file.xlsx|file.pptx]
#
# With no arguments the Tk GUI is launched. With a file argument the CLI
# extraction routine is used and text is printed to stdout.

set -euo pipefail

cd "$(dirname "$0")"

if ! command -v perl >/dev/null 2>&1; then
    cat <<'MSG' >&2
perl was not found on PATH.

Debian/Ubuntu : sudo apt install perl perl-tk libspreadsheet-parseexcel-perl
Fedora/RHEL   : sudo dnf install perl perl-Tk perl-Spreadsheet-ParseExcel
Arch          : sudo pacman -S perl perl-tk
MSG
    exit 1
fi

export PERL5LIB="$PWD:$PWD/Spreadsheet-XLSX-TempFolderCreator/lib:$PWD/Spreadsheet-XLSX-Utility2007PDP/lib:$PWD/Spreadsheet-XLSX-ZipMod2/lib:${PERL5LIB:-}"

exec perl coffice2txt.pl "$@"
