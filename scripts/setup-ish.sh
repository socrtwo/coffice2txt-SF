#!/usr/bin/env sh
# coffice2txt — iOS / iPadOS setup (a-Shell or iSH)
#
# a-Shell ships with Perl out of the box; iSH provides Alpine Linux where
# Perl is one apk away. This script detects which environment it is in and
# installs the CLI dependencies. The Tk GUI is not available on iOS.

set -eu

cd "$(dirname "$0")"

if [ -f /etc/alpine-release ]; then
    echo "==> Detected iSH (Alpine Linux)"
    apk update
    apk add --no-cache perl perl-utils perl-dev make gcc musl-dev wget curl
    if ! command -v cpanm >/dev/null 2>&1; then
        wget -qO- https://cpanmin.us | perl - --self-upgrade
    fi
    cpanm --notest --no-man-pages \
        Spreadsheet::ParseExcel \
        Spreadsheet::XLSX \
        Spreadsheet::Read || true
elif command -v perl >/dev/null 2>&1; then
    echo "==> Detected a-Shell (Perl already available)"
    echo "    a-Shell bundles Perl; CPAN modules can be installed manually if needed."
else
    echo "Unsupported environment. Install a-Shell or iSH from the App Store." >&2
    exit 1
fi

export PERL5LIB="$PWD:$PWD/Spreadsheet-XLSX-TempFolderCreator/lib:$PWD/Spreadsheet-XLSX-Utility2007PDP/lib:$PWD/Spreadsheet-XLSX-ZipMod2/lib:${PERL5LIB:-}"

cat <<MSG

iOS setup complete.

Run with:
    cd $(pwd)
    perl coffice2txt.pl path/to/file.docx > recovered.txt

Note: the Tk GUI is unavailable on iOS. Only the CLI extraction routines
are supported on this platform.
MSG
