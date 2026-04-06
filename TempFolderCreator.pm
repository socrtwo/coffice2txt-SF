package Spreadsheet::XLSX::TempFolderCreator;
use strict;
use warnings;
use 5.010000;

our $VERSION = '0.01';
use Exporter;
our @ISA = 'Exporter';
our @EXPORT = qw($temp_dir);

use File::Temp;
use base 'File::Temp::Dir'; 		# Makes temporary directory where zip 
our $temp_dir  = File::Temp->newdir;	# repaired and the extracted files will go.
