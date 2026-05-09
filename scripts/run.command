#!/usr/bin/env bash
# coffice2txt — macOS launcher
#
# Double-clickable from Finder (.command extension) or runnable from Terminal.
# Usage:
#   ./run.command [file.doc|file.xls|file.ppt|file.docx|file.xlsx|file.pptx]

set -euo pipefail

cd "$(dirname "$0")"

if ! command -v perl >/dev/null 2>&1; then
    osascript -e 'display alert "Perl not found" message "Install Perl with Homebrew: brew install perl perl-tk"' >/dev/null 2>&1 || true
    echo "perl not found. Install with: brew install perl perl-tk" >&2
    exit 1
fi

export PERL5LIB="$PWD:$PWD/Spreadsheet-XLSX-TempFolderCreator/lib:$PWD/Spreadsheet-XLSX-Utility2007PDP/lib:$PWD/Spreadsheet-XLSX-ZipMod2/lib:${PERL5LIB:-}"

exec perl coffice2txt.pl "$@"
