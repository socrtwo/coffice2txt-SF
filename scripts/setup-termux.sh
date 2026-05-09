#!/data/data/com.termux/files/usr/bin/env bash
# coffice2txt — Android (Termux) setup
#
# Installs required packages and CPAN modules, then makes the CLI runnable
# inside Termux. The Tk GUI is unavailable on Android; only the CLI text
# extraction routines are usable here.

set -euo pipefail

cd "$(dirname "$0")"

echo "==> Updating Termux packages"
pkg update -y
pkg install -y perl wget unzip make clang

echo "==> Installing CPAN modules (CLI subset, no Tk)"
if ! command -v cpanm >/dev/null 2>&1; then
    yes | cpan App::cpanminus
fi
cpanm --notest --no-man-pages \
    Spreadsheet::ParseExcel \
    Spreadsheet::XLSX \
    Spreadsheet::Read || true

export PERL5LIB="$PWD:$PWD/Spreadsheet-XLSX-TempFolderCreator/lib:$PWD/Spreadsheet-XLSX-Utility2007PDP/lib:$PWD/Spreadsheet-XLSX-ZipMod2/lib:${PERL5LIB:-}"

cat <<MSG

Termux setup complete.

Run with:
    cd $(pwd)
    perl coffice2txt.pl path/to/file.docx > recovered.txt

Note: the Tk GUI is unavailable on Android. Only the CLI extraction
routines are supported on this platform.
MSG
