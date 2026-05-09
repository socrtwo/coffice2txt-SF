# cpanfile — runtime dependencies for coffice2txt
# Install with: cpanm --installdeps .

requires 'perl', '5.010';

# Core file/path utilities (typically bundled with Perl, listed for clarity)
requires 'File::Path';
requires 'File::Basename';
requires 'File::Copy';
requires 'File::Temp';

# Spreadsheet parsing
requires 'Spreadsheet::ParseExcel';
requires 'Spreadsheet::XLSX';
requires 'Spreadsheet::Read';

# Optional GUI — only needed for the Tk front-end (coffice2txt.pl, ss2tk).
# These are listed under a feature so headless platforms (Termux, iSH) can
# skip them with: cpanm --installdeps . --without-feature=gui
feature 'gui', 'Tk graphical front-end' => sub {
    requires 'Tk';
    requires 'Tk::JComboBox';
    requires 'Tk::TableMatrix';
};

# Windows-only helpers used by the GUI launcher
on 'develop' => sub {
    recommends 'Win32';
};
