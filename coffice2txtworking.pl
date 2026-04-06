#!/usr/local/bin/perl
	# 
	#     Corrupt MS & OO Office File Salvager v. 1.0.  This program extracts text from
	#	  corrupt Microsoft and Open Office  files. It also feature format recovery
	#     for some Open Office files.
	#   
	#     Copyright (C) 2012 Paul D Pruitt
	#
	#     This program is free software: you can redistribute it and/or modify
	#     it under the terms of the GNU General Public License as published by
	#     the Free Software Foundation, either version 3 of the License, or
	#     (at your option) any later version.
	# 
	#     This program is distributed in the hope that it will be useful,
	#     but WITHOUT ANY WARRANTY; without even the implied warranty of
	#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	#     GNU General Public License for more details.
	# 
	#     You should have received a copy of the GNU General Public License
	#     along with this program.  If not, see <http://www.gnu.org/licenses/>.
	# 
	#     This program uses docx2txt project of Sandeep Kumar:
	# 	  http://docx2txt.sourceforge.net/
	# 	  And uses newbie Perl/Tk example code from:
	# 	  http://www.geocities.com/binnyva/code/perl/perl_tk_tutorial/
	# 	  It also uses InfoZip zip.exe and 7Zip's 7z.exe 
	#     which is the same I think as 7za.exe.
#

use strict;
use warnings;
use Tk;
use File::Path qw(remove_tree);
use File::Basename;
use File::Copy;
use File::Basename;
use File::Path;
use Spreadsheet::ParseExcel;
use Spreadsheet::XLSX::Fmt2007;
use Spreadsheet::XLSX::Utility2007PDP;
use Tk::DialogBox;
use Tk::FileSelect;
use Tk::JComboBox;
use Tk::Menu;
use Tk::Menubutton;
use Tk::NoteBook;
use Tk::Pretty;
use Tk::Button;
use Tk::TableMatrix::Spreadsheet;
BEGIN {
	Win32::SetChildShowWindow(0)
	if defined &Win32::SetChildShowWindow
};

# 
	# Create the Main Window
# 

my $mw = new MainWindow;
$mw->title('Corrupt Office Salvager');

# 
	# Hides TK logo with my own logo
#

my $icon = $mw->Photo(-file => 'icon_32x32.gif');
$mw->iconimage($icon);

# 
	# Declare that there is a menu, create text 
	# editor and create a vertical scroll bar
#

my $mbar = $mw -> Menu();
$mw -> configure(-menu => $mbar);
my $textarea = $mw -> Frame(); #Creating Another Frame
my $txt = $textarea -> Text(-width=>80, -height=>22);
my $srl_y = $textarea -> Scrollbar(-orient=>'v',-command=>[yview => $txt]);
$txt -> configure(-yscrollcommand=>['set', $srl_y]);
$txt -> grid(-row=>1,-column=>1);
$srl_y -> grid(-row=>1,-column=>2,-sticky=>"ns");
$textarea -> grid(-row=>5,-column=>1,-columnspan=>2);

# 
	# Main Menu choices setup section.
# 

my $file = $mbar -> cascade(-label=>"File", -underline=>0, -tearoff => 0);
my $alternatives = $mbar -> cascade(-label=>"Alternatives", -tearoff => 0);
my $help = $mbar -> cascade(-label =>"Help", -underline=>0, -tearoff => 0);

# 
	# File Menu choices setup section.
# 

$file -> command(-label =>"Extract I", -underline => 0,
-command => [\&menuOpenClickedExtract1, "Open"]);
$file -> command(-label =>"Extract II", -underline => 0,
-command => [\&menuOpenClickedExtract2, "Open"]);
$file -> command(-label =>"Full Recovery \(Open Office Only\)", -underline => 0,
-command => [\&openOfficeFormatRecovered, "Open"]);
$file -> command(-label =>"Save", -underline => 0,
-command => [\&menuSavedClicked, "Save"]);
$file -> separator();
$file -> command(-label =>"Exit", -underline => 1,
-command => sub { exit } );

# 
	# Alternatives menu choices setup section.
# 		
my $wordalternatives = $alternatives -> cascade(-label =>"Word Recovery", -underline => 0, -tearoff => 0);
$wordalternatives -> command(-label =>"Steps to Recovering a Corrupt Word File - free doc and docx recovery advice", -command => [\&wordStepsClicked, "Steps to Recovering a Corrupt Word File - free doc and docx recovery advice"]);
$wordalternatives -> command(-label =>"Repair My Word - freeware doc document text extractor", -command => [\&repairMyWordClicked, "Repair My Word - freeware doc document text extractor"]);
$wordalternatives -> command(-label =>"Corrupt DOCX Salvager - freeware for docx recovery", -command => [\&corruptDocxSalvagerClicked, "Corrupt DOCX Salvager - freeware for docx recovery"]);	
$wordalternatives -> command(-label =>"Savvy DOCX Recovery - freeware for docx text and format recovery", -command => [\&savvyWordRecoveryClicked, "Savvy DOCX Recovery - freeware for docx text and format recovery"]);
my $excelalternatives = $alternatives -> cascade(-label =>"Excel Recovery", -underline => 0, -tearoff => 0);
$excelalternatives -> command(-label =>"Steps to Recovering a Corrupt Excel File - free xls and xlsx recovery advice", -command => [\&excelStepsClicked, "Steps to Recovering a Corrupt Excel File - free xls and xlsx recovery advice"]);
$excelalternatives -> command(-label =>"Excel Recovery - freeware for xls and xlsx recovery", -command => [\&excelRecoveryClicked, "Excel Recovery - freeware for xls and xlsx recovery"]);
$excelalternatives -> command(-label =>"Corrupt XLSX Salvager - freeware for xlsx recovery", -command => [\&corruptXlsxSalvagerClicked, "Corrupt XLSX Salvager - freeware for xlsx recovery"]);
$wordalternatives -> command(-label =>"Corrupt MS Office 2007/2010 Extractor - docx, xlsx and pptx repair freeware", -command => [\&cptOfficeRecoveryClicked, "Corrupt MS Office 2007/2010 Extractor - docx, xlsx and pptx repair freeware"]);	
$wordalternatives -> command(-label =>"Recovery for Word - commercial software for doc and docx recovery", -command => [\&recoveryForWordClicked, "Recovery for Word - commercial software for doc and docx recovery"]);	
$wordalternatives -> command(-label =>"Repair Word Documents Online - commercial doc and docx service", -command => [\&recoveronixWordServiceClicked, "Repair Word Document Online - commercial doc and docx service"]);		
$wordalternatives -> command(-label =>"WordFix - commercial software for doc and docx recovery", -command => [\&wordFixClicked, "WordFix - commercial software for doc and docx recovery"]);
$excelalternatives -> command(-label =>"Recovery for Excel - commercial software for xls and xlsx recovery", -command => [\&recoveryForExcelClicked, "Recovery for Excel - commercial software for xls and xlsx recovery"]);	
$excelalternatives -> command(-label =>"Repair Excel Documents Online - commercial xls and xlsx service", -command => [\&recoveronixExcelServiceClicked, "Repair Excel Documents Online - commercial xls and xlsx service"]);		
$excelalternatives -> command(-label =>"ExcelFix - commercial software for xls and xlsx recovery", -command => [\&excelFixClicked, "ExcelFix - commercial software for xls and xlsx recovery"]);
$excelalternatives -> command(-label =>"Corrupt MS Office 2007/2010 Extractor - docx, xlsx and pptx repair freeware", -command => [\&cptOfficeRecoveryClicked, "Corrupt MS Office 2007/2010 Extractor - docx, xlsx and pptx repair freeware"]);	
my $powerpointalternatives = $alternatives -> cascade(-label =>"PowerPoint Recovery", -underline => 0, -tearoff => 0);
$powerpointalternatives -> command(-label =>"Recovery for PowerPoint - commercial software for ppt and pptx recovery", -command => [\&recoveryForPowerPointClicked, "Recovery for PowerPoint - commercial software for ppt and pptx recovery"]);	
$powerpointalternatives -> command(-label =>"Repair PowerPoint Documents Online - commercial ppt and pptx service", -command => [\&recoveronixPowerPointServiceClicked, "Repair PowerPoint Documents Online - commercial ppt and pptx service"]);
$powerpointalternatives -> command(-label =>"Corrupt MS Office 2007/2010 Extractor - docx, xlsx and pptx repair freeware", -command => [\&cptOfficeRecoveryClicked, "Corrupt MS Office 2007/2010 Extractor - docx, xlsx and pptx repair freeware"]);
my $openofficealternatives = $alternatives -> cascade(-label =>"Open Office Recovery", -underline => 0, -tearoff => 0);
$openofficealternatives -> command(-label =>"10 Free Ways to Recover Open Office - free odt, ods and odp recovery advice", -command => [\&openOfficeStepsClicked, "10 Free Ways to Recover Open Office - free odt, ods and odp recovery advice"]);
$openofficealternatives -> command(-label =>"Corrupt Open Office Recovery - odt, ods and odp recovery freeware", -command => [\&corruptOpenOfficeRecoveryClicked, "Corrupt Open Office Recovery - odt, ods and odp recovery freeware"]);
$openofficealternatives -> command(-label =>"Recovery for Open Office Writer - commercial software for odt recovery", -command => [\&recoveryForOpenOfficeWriterClicked, "Recovery for Open Office Writer - commercial software for odt recovery"]);
$openofficealternatives -> command(-label =>"Repair Open Office Writer Online - commercial odt service", -command => [\&recoveronixOpenOfficeWriterServiceClicked, "Repair Open Office Writer Online - commercial odt service"]);
$openofficealternatives -> command(-label =>"Recovery for Open Office Calc - commercial software for ods recovery", -command => [\&recoveryForOpenOfficeCalcClicked, "Recovery for Open Office Writer - commercial software for odt recovery"]);
$openofficealternatives -> command(-label =>"Repair Open Office Calc Online - commercial ods service", -command => [\&recoveronixOpenOfficeCalcServiceClicked, "Repair Open Office Calc Online - commercial ods service"]);
$alternatives -> command(-label =>"Contact Me - upload corrupt files for manual recovery", -command => [\&sendToMeClicked, "Contact Me - upload corrupt docx document for manual recovery"]);

# 
	# Help Menu choices setup section. 
# 

$help -> command(-label =>"About and Instructions", -command => sub { 
	$txt->delete('1.0','end');
	$txt->insert('end',
	"*****Corrupt Office Salvager About and Instructions******
	
	-----How to use this program------
	
	1.  Click on the File Menu and choose Full Recovery or one of the two
	salvage methods.
	2.  Choose your Microsoft or Open Office file.
	3.  If you chose a salvage method, your extracted text will be displayed.
	If you chose a the Full Recovery choice the application will attempt
	to launch a recovered version with the software currently assigned as
	the default for the corrupt file's extension, odt, ods or odp.
	4.  If you chose a Salvage method, you can Edit, the text as desired.
	5.  With salvaged text, choose the Save menu choice on the File Menu and 
	file to the name and file location you wish.
	
	-----About-----
	
	This program will extract the text from some corrupted or all healthy 
	Microsoft Office and Open Office files (2.X and 3.X files) with the 
	extensions .doc, docx, xls, xlsx, ppt, pptx, odt, ods and odp as well 
	as possibly the template and macro variants of these extensions such as 
	dot, xlt and pps if they are changed to the correct corresponding 
	extensions mentioned. It may succeed at doing so where MS Office Open 
	Office itself fails to salvage text. It can also attempt to recover 
	formatting in the form of a full Open Office file with a regular, odt, 
	ods or odp extension. At this time there is no facility for recovering 
	anything but basic formatting for MS Office files through the previously 
	mentioned text extractions although links are provided on the Alternatives 
	menu. This program can be used as a viewer of text within healthy 
	MS Office and Open Office files without having Open Office installed. 
	
	The text extraction is accomplished with the use of the command line
	application, SILVERCODERS DocToText. The program also uses command line tools 
	from The Chicago Project extract data and text from Excel and PowerPoint 
	97-2003 format files.  The reconstructed version of the Open Office file 
	is accomplished by unzipping the Open Office file with the somewhat zip 
	corruption immune 7z.exe unzipper. Once unzipped, the manifest/manifest.xml
	file is replaced with a greatly simplified version as described here: http://
	www.oooforum.org/forum/viewtopic.phtml?t=57600. 
	
	If this application doesn't work, there are other things worth trying 
	as summarized here: http://s2services.com/open_office.htm
	
	----Changes to 1.0----
	
	1. Added a zip repair pretreatment using InfoZip zip.exe -FF command for all zip 
	based file recoveries.
	2. New icon.
	3. Using 7zip as the the archive extractor now instead of CakeCMD or No Frills
	unzipper.
	4. Preserved the do_not_remove folder with its manifest.xml contents during the 
	installation routine. This issue and the next one originally prevented the corrupt 
	Open Office recovery feature working in version 0.22.
	5. Also fixed the preserving of the unzipped folder and its sevenzipcmd.exe 
	contents as well as the support dll files in the installation routine.
	
	-----Credits-----
	
	This program is made by Paul D Pruitt (socrtwo) and uses the following
	command line applications in its operation: SILVERCODERS DocToText; 
	xlhtml and ppthml from The Chicago project; Runar Skaret's 
	ReadText; cakecmd.exe unzipper by Leung Yat Chun Joseph; No-Frills Unzipper 
	by Ccy; 7-Zip CMD rezipper; and Nirsoft's HTMLAsText. It also uses Perl/Tk 
	code for the GUI elements as described here http://www.bin-co.com/perl/
	perl_tk_tutorial/. 
	
	Here are the links:
	* ReadText:http://members.fortunecity.com/bigg5/frw/diagn.htm
	* DocToText: http://silvercoders.com/en/products/doctotext
	* Xlhtml and Ppthtml: http://prdownloads.sf.net/chicago/xlHtml-Win32-040.zip
	* No-Frills Unzipper: http://godskingsandheroes.info/software/
	#no-frills_command_line_unzipper 
	* CakeCMD Unzipper:http://www.quickzip.org/softwares-cakecmd and for .Net. 2.0, see: 
	http://filehippo.com/download_dotnet_framework_2/
	* 7-Zip Command Line Version: http://www.7-zip.org/download.html
	* NirSoft's HtmlAsText: http://www.nirsoft.net/utils/htmlastext.html
	
	-----Contact Info-----
	* My software website is http://www.godskingsandheroes.info/software/.
	* Also visit my data recovery software list http://www.s2services.com.
	* My E-Mail: socrtwo\@s2services
	* My phone number is 301-493-4982. 
	* I do data recovery for \$22 an incident. I sometimes do charity work.
	"); 
});

# 
	# Open and saved dialog box scalar declarations.
# 

my $typesOpen = [ ['Microsoft and Open Office Document', '.doc .docx .xls .xlsx .ppt .pptx .odt .ods .odp'],['All files', '*'],];
my $typesOpen2 = [ ['Open Office Document', '.odt .ods .odp'],['All files', '*'],];
my $typesSaved = [ ['Text files', '.txt'], ['All files',   '*'],];

# 
	# Most of the scalars and the one array used declarations.
# 

my $lowerCaseFileName;
my ($fullName, $dirName, $mainFilePath, $name, $path, $suffix, $but, $lowerCaseDirectoryName, $newLongPath);
my ($fileToBeCopied, $newManifested, $wfh, $salvagedText, $dir, $lowerCaseBaseName, $underScoreLowerCaseBaseName);
my ($objFSO, $zipFileName, $content, $saved, $docx_name, $editedContent, $recoveredFile, $cpRecoveredFile, $strFolderPath, );
my @suffixlist;
my $nl = "\r\n";                # Alternative is "\n".
my $lineIndent = "  ";     # Indent nested lists by "\t", " " etc. I don't think I use this scalar.
my $lineWidth = 80;        # Line width, used for short line justification. I don't think I use this scalar either.

# 
	# I think this is the deceleration of the file used to capture errors.
# 

my $processing = "processing.txt";	

# 
	# Main loop currently activated by selecting the file.
# 

MainLoop;
sub menuOpenClickedExtract1 {
	
	# 
		# Code displaying the File Open Dialog Box and 
		# records the file choice in the $mainFilePath scalar.
	# 
	
	$mainFilePath = $mw->getOpenFile(-filetypes => $typesOpen, -defaultextension => '.doc .docx .xls .xlsx .ppt .pptx .odt .ods .odp');	  
	return unless $mainFilePath;
	
	# 
		# Change $mainFilePath to lower case.
	# 
	
	$lowerCaseFileName = lc($mainFilePath);
	$lowerCaseBaseName = basename($lowerCaseFileName, @suffixlist);
	$lowerCaseDirectoryName = dirname($lowerCaseFileName);
	print "$lowerCaseBaseName$nl";
	
	# 
		# Replace $lowerCaseFileName spaces with underscores.
	# 	
	
	$underScoreLowerCaseBaseName = $lowerCaseBaseName;
	$underScoreLowerCaseBaseName =~ s/ /_/g;
	my $usNamedUnderScoreLowerCaseBaseName = 'us_' . $underScoreLowerCaseBaseName;
	copy($mainFilePath, $usNamedUnderScoreLowerCaseBaseName) or warn "Unable to copy $mainFilePath to $usNamedUnderScoreLowerCaseBaseName.$nl";
	print "$usNamedUnderScoreLowerCaseBaseName$nl";
	
	# 
		# For some reason that I don't remember, maybe because this was originally an online   
		#  script, I don't work with, the filename but rename it to a random number instead.
	# 	
	
	my @file_type   = split(/\./, $underScoreLowerCaseBaseName);
	my $file_type   = $file_type[$#file_type];
	my $random_long = int(rand(10000000));
	$salvagedText = $random_long.'.txt';
	print "$salvagedText$nl";
	
	my $randomName = $random_long . "." . $file_type;
	copy($usNamedUnderScoreLowerCaseBaseName, $randomName) or warn "Unable to copy $underScoreLowerCaseBaseName to $randomName.$nl";
	print "$randomName$nl";
	
	#
		# Here we start processing of the file chosen in the Extract I 
		# file dialog box. In this case if we chose a ppt file.
	#
	
	if($file_type eq 'ppt'){	
		
		#
			# Trying to recover text with rt.exe actually doesn't work on 64 
			# bit machines. I will have to find a replacement text extractor.
		#	
		
		open $wfh, "| coffec.exe -t $randomName > $salvagedText 2> $processing" or warn "Unable to use rt to extract text from $randomName to $salvagedText.$nl";
		close $wfh;
		
		#
			# Here we read the resulting salvaged text into the $_ scalar. 
		#	
		
		{
			local $/=undef;
			open FILE, "$salvagedText" or warn "Couldn't open $salvagedText for writing into the scalar which is in turn written into the text area.$nl";
			binmode FILE;
			$_= <FILE>;
			close FILE;
		}
		
		#
			# Here we just delete any text previously in the 
			# text box and substitute with the salvaged text.
		#	
		
		$txt-> delete('1.0','end');
		$txt -> insert('end',$_ );
		
		#
			# Here we delete temporary files we  
			# made that are no longer needed.
		#	
		
		unlink $randomName;
		unlink $usNamedUnderScoreLowerCaseBaseName;
		unlink $salvagedText;		
		
		#
			# Next we process doc and xls  
			# extensioned files with doctotext.
		#		
		
		} elsif ($file_type eq 'doc' or $file_type eq 'xls'){	
		
		#
			# For any extension that has an underlying zip structure, doctotext 
			# needs more complicated commands, that is not true with doc and xls.
		#		
		
		open $wfh, "| doctotext.exe $randomName > $salvagedText 2> $processing" or warn "Unable to use DocToText to extract data from $randomName to $salvagedText.$nl";
		close $wfh;
		
		#
			# Reading the resulting salvaged text into the $_ scalar. 
		#	
		
		{
			local $/=undef;
			open FILE, "$salvagedText" or warn "Couldn't open $salvagedText for writing into the scalar which is in turn written into the text area.$nl";
			binmode FILE;
			$_= <FILE>;
			close FILE;
		}
		
		#
			# Deleting previously displayed text in the  
			# text box and substituting the salvaged text.
		#	
		
		$txt-> delete('1.0','end');
		$txt -> insert('end',$_ );
		
		#
			# Deleting temporary files.
		#	
		
		unlink $randomName;
		unlink $usNamedUnderScoreLowerCaseBaseName;
		unlink $salvagedText;
		
		#
			# Now we process all the files with zip structure. Those  
			# are the extensions docx, xlsx, pptx, odt, ods and odp.
		#	
		
 		} else {
		
		#
			# Creates temporary folder with help of Spreadsheet::XLSX::TempFolderCreator.
		#
		
		use Spreadsheet::XLSX::TempFolderCreator;
		print "$nl My temporary directory from xlsx2csv's script point of view is: $temp_dir$nl";
		my $randomNamedZipRepairedFilePath  =  $temp_dir . "\\zip_repaired_" . $randomName; 
		
		#
			# Repairs our randomly renamed $mainFilePath zip structured 
			# file and saves the results in a temporary folder.
		#
		
		system ("cmd /c zip.exe -FF \"$randomName\" --out \"$randomNamedZipRepairedFilePath\" ");	
		
		#
			# Uses DocToText to recover the text from a zip structure MS or Open Office file.
			# Uses 7zips 7z.exe command line version to help. %a means where the archive 
			# normally sits in a 7z.exe command line. %f stands for a specific file requested
			# to be extracted from the archive and %d stands for directory. A normal sample 
			# 7z.exe extraction would look something like this: 
			# C:\>7z.exe x myCorruptDocument.docx.zip document.xml -o"C:\Temp". So there is normally
			# not a space between the -o switch and the output directory it is setting. So
			# DocToText needs to be told how 7zip operates in terms of the commands witches and the
			# order of the archive, target file to be extracted and output directory.
		#
		
		open $wfh, "| doctotext.exe --fix-xml --unzip-cmd=\"7z.exe x %a %f -y -o%d\" \"$randomNamedZipRepairedFilePath\" > $salvagedText 2> $processing" or warn "Unable to use DocToText to extract text from $randomName to $salvagedText.$nl";
		close $wfh;
		
		#
			# Reading the resulting salvaged text into the $_ scalar. 
		#	
		
		{
			local $/=undef;
			open FILE, "$salvagedText" or warn "Couldn't open $salvagedText for writing into the scalar which is in turn written into the text area.$nl";
			binmode FILE;
			$_= <FILE>;
			close FILE;
		}
		
		#
			# Deleting previously displayed text in the  
			# text box and substituting the salvaged text.
		#	
		
		$txt-> delete('1.0','end');
		$txt -> insert('end',$_ );
		
		#
			# Deleting temporary files.
		#	
		
		unlink $randomNamedZipRepairedFilePath or warn "Could not unlink $randomNamedZipRepairedFilePath";
		unlink $randomName;
		unlink $usNamedUnderScoreLowerCaseBaseName;
		unlink $salvagedText;
		
	}}
	
	#
		# This is the code for the "File Extract II" menu choice.
	#	
	
	sub menuOpenClickedExtract2 {
		
		# 
			# Code displaying the File Open Dialog Box which
			# records the file choice in the $mainFilePath scalar.
		# 
		
		my $mainFilePath2 = $mw->getOpenFile(-filetypes => $typesOpen, -defaultextension => '.doc .docx .xls .xlsx .ppt .pptx .odt .ods .odp');	  
		return unless $mainFilePath2;	
		
		# 
			# Change $mainFilePath to lower case.
		# 
		
		my $lowerCaseFileName2 = lc($mainFilePath2);
		my $lowerCaseBaseName2 = basename($lowerCaseFileName2,@suffixlist);
		my $lowerCaseDirectoryName2 = dirname($lowerCaseFileName2);
		print "$lowerCaseBaseName2$nl";
		
		# 
			# Replace $lowerCaseFileName spaces with underscores.
		# 	
		
		$lowerCaseBaseName2 =~ s/ /_/g;
		my $usNamedUnderScoreLowerCaseBaseName2 = 'us_' . $lowerCaseBaseName2;
		copy($mainFilePath2, $usNamedUnderScoreLowerCaseBaseName2) or warn "Unable to copy $mainFilePath to $usNamedUnderScoreLowerCaseBaseName2.$nl";
		print "$usNamedUnderScoreLowerCaseBaseName2$nl";
		
		# 
			# For some reason that I don't remember, maybe because this was originally an online   
			#  script, I don't work with, the filename but rename it to a random number instead.
		# 	
		
		my @file_type2 = split(/\./, $lowerCaseBaseName2);
		my $file_type2 = $file_type2[$#file_type2];
		my $random_long2 = int(rand(10000000));
		my $salvagedhtml = $random_long2.'.html';
		print "$salvagedhtml$nl";
		my $salvagedText2 = $random_long2.'.txt';
		print "$salvagedText2$nl";
		
		my $randomName2 = $random_long2 . "." . $file_type2;
		copy($usNamedUnderScoreLowerCaseBaseName2, $randomName2) or warn "Unable to copy $usNamedUnderScoreLowerCaseBaseName2 to $randomName2.$nl";
		print "$randomName2$nl";
		
		#
			# Here we start processing of the file chosen in the Extract 
			# II file dialog box. In this case if we chose a ppt file.
		#	
		
		if($file_type2 eq 'ppt'){
			
			#
				# Ppthtml.exe extracts and converts the ppt to HTML.
			#	
			
			open $wfh, "| ppthtml.exe $randomName2 > $salvagedhtml 2> $processing" or warn "Unable to use ppthtml to convert the ppt file from $randomName2.$nl";
			close $wfh;
			
			#
				# I don't know what these lines do except maybe create an 
				# html document. I don't know where I got the code either...
			#	
			
			use Cwd 'abs_path';
			my $cfgabspath = dirname(abs_path($0)).'\test.cfg';
			my $relhtmldir = dirname(abs_path($0)).'\text.html';
			my $reltextdir = dirname(abs_path($0)).'\text.txt';
			print "$cfgabspath$nl";
			print "$relhtmldir$nl";
			print "$reltextdir$nl";
			
			open FH, ">test.cfg";
			print FH "[Config]\n";
			print FH "OpenInNotepad=0\n";
			print FH "CharsPerLine=80\n";
			print FH "Source=$relhtmldir\n";
			print FH "Dest=$reltextdir\n";
			print FH "SkipTitleText=0\n";
			print FH "AddLineUnderHeader=0\n";
			print FH "SkipTableHeaderText=0\n";
			print FH "TableCellDelimit=2\n";
			print FH "HeadingLineChars=======\n";
			print FH "HorRuleChar==\n";
			print FH "ListChars=*o-@#\n";
			print FH "ConvertMode=1\n";
			print FH "AllowCenterText=0\n";
			print FH "AllowRightText=0\n";
			print FH "DLSpc=8\n";
			print FH "LinksDisplayFormat=%T\n";
			print FH "EncloseBoldCharsStart=<<\n";
			print FH "EncloseBoldCharsEnd=>>\n";
			print FH "EncloseBold=0\n";
			print FH "SubFolders=0\n";
			print FH "\n";
			close FH; 
			
			#
				#  I guess this code extracts text from the html file.
			#
			
			my $texthtmlname = "text.html";
			copy($salvagedhtml, $texthtmlname) or warn "Unable to copy $salvagedhtml to text.html.$nl";
			print "$texthtmlname$nl";
			my $textname = "text.txt";
			open $wfh, "| HtmlAsText.exe /run \"$cfgabspath\" 2> $processing" or warn "Unable to use HtmlAsText.exe to extract text from the html results file produced by ppthtml.$nl";
			close $wfh;
			
			#
				# Reading the resulting salvaged text into the $_ scalar. 
			#	
			
			{
				local $/=undef;
				open FILE, "$textname" or warn "Couldn't open $textname for writing into the scalar which is in turn written into the text area.$nl";
				binmode FILE;
				$_= <FILE>;
				close FILE;
			}
			
			#
				# Deleting previously displayed text in the  
				# text box and substituting the salvaged text.
			#	
			
			$txt-> delete('1.0','end');
			$txt -> insert('end',$_ );
			
			#
				# Deleting temporary files.
			#	
			
			unlink $randomName2;
			unlink $usNamedUnderScoreLowerCaseBaseName2;
			unlink $salvagedhtml;
			unlink $texthtmlname;
			unlink $textname;
			
			#
				# Processing of the file chosen in the Extract 
				# II file dialog box if its an xls spreadsheet.
			#	
			
			} elsif ($file_type2 eq 'xlsx') {
			
			use Cwd;
			my $originaldir = getcwd;
			print "$nl My current working directory is $originaldir$nl";
			
			#
				# Creates temporary folder with help of Spreadsheet::XLSX::TempFolderCreator.
			#
			
			use Spreadsheet::XLSX::TempFolderCreator;
			print "$nl My temporary directory from xlsx2csv's script point of view is: $temp_dir$nl";
			my $randomNamedZipRepairedFilePath2  =  $temp_dir . "\\zip_repaired_" . $randomName2; 
			
			#
				# Repairs our randomly renamed $mainFilePath zip structured 
				# file and saves the results in a temporary folder.
			#
			
			system ("cmd /c zip.exe -FF \"$randomName2\" --out \"$randomNamedZipRepairedFilePath2\" ");	
			
			#
				# This alternate method for method for processing MS Office 2007/2010 format files uses CMD Corrupt 
				# OfficeOpen2txt which in turn uses no-frills unzipper and XML Partner to extract text/data.
			#
			
			open $wfh, "| coffec.exe -t \"$randomNamedZipRepairedFilePath2\" " or warn "Unable to convert $randomNamedZipRepairedFilePath2 to CSV files$nl";
			close $wfh;
			
			chdir $temp_dir;
			use File::Copy;
			for ( glob '*.csv' ) {
				move $_, $lowerCaseDirectoryName2;
			}
			
			#
				# Deleting previously displayed text in the  
				# text box and substituting the salvaged text.
			#	
			
			$txt-> delete('1.0','end');
			$txt -> insert('end','

			If successful this mode will
			extract	the recoverable worksheets
			and their data into csv files named
			with the worksheet names and saved 
			in the same folder from where you 
			chose your corrupted file.');
			
			my $xldirectory = $temp_dir . "\\xl";
			
			chdir $originaldir;
			unlink $randomNamedZipRepairedFilePath2 or warn "$nl Could not unlink $randomNamedZipRepairedFilePath2.$nl";
			unlink $randomName2 or warn "$nl Could not unlink $randomName2.$nl";
			unlink $usNamedUnderScoreLowerCaseBaseName2 or warn "$nl Could not unlink $usNamedUnderScoreLowerCaseBaseName2.$nl";
			remove_tree($xldirectory) or warn "Cannot remove $xldirectory directory.$nl";
			
			#
				# Processes Word and PowerPoint 2007/2010 format files.
			#
			
			} elsif ($file_type2 eq 'docx' or $file_type2 eq 'pptx') {
			
			#
				# Creates temporary folder with help of Spreadsheet::XLSX::TempFolderCreator.
			#
			
			use Spreadsheet::XLSX::TempFolderCreator;
			print "$nl My temporary directory from xlsx2csv's script point of view is: $temp_dir$nl";
			my $randomNamedZipRepairedFilePath2  =  $temp_dir . "\\zip_repaired_" . $randomName2; 
			
			#
				# Repairs our randomly renamed $mainFilePath zip structured 
				# file and saves the results in a temporary folder.
			#
			
			system ("cmd /c zip.exe -FF \"$randomName2\" --out \"$randomNamedZipRepairedFilePath2\" ");	
			
			#
				# This alternate method for method for processing MS Office 2007/2010 format files uses CMD Corrupt 
				# OfficeOpen2txt which in turn uses no-frills unzipper and XML Partner to extract text/data.
			#
			
			open $wfh, "| coffec.exe -t $randomNamedZipRepairedFilePath2 > $salvagedText2 2> $processing" or warn "Unable to use rt to extract text from $randomName2 to $salvagedText2.$nl";
			close $wfh;
			
			rename $randomNamedZipRepairedFilePath2, $randomName2;
			
			#
				# Reading the resulting salvaged text into the $_ scalar. 
			#	
			
			local $/=undef;
			open FILE, "$salvagedText2" or warn "Couldn't open $salvagedText for writing into the scalar which is in turn written into the text area.$nl";
			binmode FILE;
			$_= <FILE>;
			close FILE;
			
			#
				# Deleting previously displayed text in the  
				# text box and substituting the salvaged text.
			#	
			
			$txt-> delete('1.0','end');
			$txt -> insert('end',$_ );
			
			#
				# Deleting temporary files.
			#	
			
			unlink $randomNamedZipRepairedFilePath2;
			unlink $randomName2;
			unlink $usNamedUnderScoreLowerCaseBaseName2;
			unlink $salvagedText2;
			
			#
				# Prcesses doc, odt, ods and odp files, the 
				# latter three being Open Office extensions.
			#
			
			} else {
			
			#
				# Creates temporary folder with help of Spreadsheet::XLSX::TempFolderCreator.
			#
			
			use Spreadsheet::XLSX::TempFolderCreator;
			print "$nl My temporary directory from xlsx2csv's script point of view is: $temp_dir$nl";
			my $randomNamedZipRepairedFilePath2  =  $temp_dir . "\\zip_repaired_" . $randomName2; 
			
			#
				# Repairs our randomly renamed $mainFilePath zip structured 
				# file and saves the results in a temporary folder.
			#
			
			system ("cmd /c zip.exe -FF \"$randomName2\" --out \"$randomNamedZipRepairedFilePath2\" ");	
			
			#
				# Uses DocToText to recover the text from a zip structured Open Office file.
				# Uses 7zips 7z.exe command line version to help. %a means where the archive 
				# normally sits in a 7z.exe command line. %f stands for a specific file requested
				# to be extracted from the archive and %d stands for directory. A normal sample 
				# 7z.exe extraction would look something like this: 
				#
				# C:\>7z.exe x myCorruptDocument.docx.zip document.xml -o"C:\Temp" 
				#
				# So there is normally not a space between the -o switch and the output directory it 
				# is setting.So DocToText needs to be told how 7zip operates in terms of the commands 
				# switches and the order of the archive, target file to be extracted and output directory.
				#
				# The difference between this extraction and the the one in Extraction I is the use of 
				# the --strip-xml switch instead of the --fix-xml one with presumably more raw text here 
				# and more formatted results with the other method. Also only Open Office files are 
				# processed here. 
			#
			
			open $wfh, "| doctotext.exe --strip-xml --unzip-cmd=\"7z.exe x %a %f -y -o%d\" \"$randomNamedZipRepairedFilePath2\" > $salvagedText2 2> $processing" or warn "Unable to use DocToText to extract text from $randomNamedZipRepairedFilePath2 to $salvagedText2.$nl";
			close $wfh;
			rename $randomNamedZipRepairedFilePath2, $randomName2;
			
			#
				# Reading the resulting salvaged text into the $_ scalar. 
			#	
			
			{
				local $/=undef;
				open FILE, "$salvagedText2" or warn "Couldn't open $salvagedText for writing into the scalar which is in turn written into the text area.$nl";
				binmode FILE;
				$_= <FILE>;
				close FILE;
			}
			
			#
				# Deleting previously displayed text in the  
				# text box and substituting the salvaged text.
			#	
			
			$txt-> delete('1.0','end');
			$txt -> insert('end',$_ );
			
			#
				# Deleting temporary files.
			#	
			
			unlink $randomNamedZipRepairedFilePath2;
			unlink $randomName2;
			unlink $usNamedUnderScoreLowerCaseBaseName2;
			unlink $salvagedText2;
			
		}}
		
		#
			# Code activated by the Save choice on the File Menu.
		#
		
		sub menuSavedClicked {
			
			#
				# Gets the results from the visible Tk text 
				# box Window which can even have been edited.
			# 
			
			$editedContent = $txt->Contents();
			
			#
				# Gets the file name from the choice in the Save File Dialog Box.
			#
			
			$saved = $mw->getSaveFile(-filetypes => $typesSaved,
			-defaultextension => '.txt');
			return unless $saved;
			
			#
				# Takes text from the scalar containing the edited contents and
				# saves it to the chosen file name.
			#
			
			open (MYFILE, ">$saved");
			print MYFILE "$editedContent";
			close (MYFILE);
			close $saved;
		}
		
		#
			# This code is activated when choosing the 
			# Open Office file recovery File menu option.
		#
		
		sub openOfficeFormatRecovered {
			
			#
				# Opens a file or specifically allows choice
				# of an Open Office file name to work on.
			#
			
			my $mainFilePath3 = $mw->getOpenFile(-filetypes => $typesOpen2, -defaultextension => '.odt, ods, .odp');	  
			return unless $mainFilePath3;
			
			#
				# Changes the chosen file name and its complete path to lower case.
			#
			
			my $lowerCaseFileName3 = lc($mainFilePath3);
			my $lowerCaseBaseName3 = basename($lowerCaseFileName3,@suffixlist);
			my $lowerCaseDirectoryName3 = dirname($lowerCaseFileName3);
			print $lowerCaseBaseName3;
			
			#
				# Replaces spaces in the base name with underscores.
			#
			use Spreadsheet::XLSX::TempFolderCreator;
			print "$nl My temporary directory from xlsx2csv's script point of view is: $temp_dir$nl";
			
			my $underScoreLowerCaseBaseName3 = $lowerCaseBaseName3;
			$underScoreLowerCaseBaseName3 =~ s/ /_/g;
			my $usNamedUnderScoreLowerCaseBaseName3 = $temp_dir . "\\us_" . $underScoreLowerCaseBaseName3;
			
			copy($mainFilePath3, $usNamedUnderScoreLowerCaseBaseName3) or warn "unable to copy $mainFilePath to $usNamedUnderScoreLowerCaseBaseName3.$nl";
			
			#
				# Repairs a zip extensioned version of the chosen 
				# corrupt Open Office file to a temporary folder.
			#
			
			
			my $zipRepairedFilePath  =  $temp_dir . "\\zip_Repaired3_" . $underScoreLowerCaseBaseName3; 
			
			system ("cmd /c zip.exe -FF \"$usNamedUnderScoreLowerCaseBaseName3\" --out \"$zipRepairedFilePath\" ");
			
			#
				# unzipping of the zip repaired Open Office Text file.
			#
			
			open $wfh, "| 7z.exe x \"$zipRepairedFilePath\" -y -ounzipped " or warn "Unable to unzip $zipFileName with 7zip.$nl";
			close $wfh;
			
			
			#
				# manifest.xml is replaced by a copy with most of the 
				# style XML removed as described in Open Office forums.
			#
			
			use File::Copy;
			$fileToBeCopied = 'do_not_remove\manifest.xml';
			$newManifested = 'unzipped\META-INF\manifest.xml';
			copy($fileToBeCopied, $newManifested) or warn "Copy of do_not_remove\\manifest.xml to unzipped\\META-INF\\manifest.xml failed.$nl";
			chdir('unzipped') or warn "Cannot change directory to unzipped.$nl";
			$recoveredFile = 'recovered_' . $underScoreLowerCaseBaseName3;
			
			#
				# Extracted files with hopefully "repaired" mainfest.xml file are rezipped.
			#
			
			open $wfh, "| ..\\sevenzipcmd.exe a -tzip $recoveredFile \* " or warn "Unable to rezip $recoveredFile.$nl";
			close $wfh;
			
			#
				# Various hijinks to move corrected file to 
				# directory where we started and remove temporary files.
			#
			
			$cpRecoveredFile = $lowerCaseDirectoryName3 . "\\" . $recoveredFile;
			move($recoveredFile, $cpRecoveredFile) or warn "File cannot move $recoveredFile to $cpRecoveredFile.$nl";
			chdir('../') or warn "Cannot change directory back to root.$nl";
			remove_tree('unzipped') or warn "Cannot remove unzipped directory.$nl";
			unlink $usNamedUnderScoreLowerCaseBaseName3;
			unlink $zipRepairedFilePath;
			system "$cpRecoveredFile";
		}
		
		#
			# Web page navigation subroutine activated by 
			# clicking on the Alternatives menu choices.
		#	
		
		sub savvyWordRecoveryClicked {
			
			my $url = "http://www.godskingsandheroes.info/software/#17";
			my $commandline = qq{start "$url" "$url"}; 	
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub repairMyWordClicked {
			
			my $url = "http://www.repairmyword.com/";
			my $commandline = qq{start "$url" "$url"}; 	
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub wordStepsClicked {
			
			my $url = "http://www.s2services.com/word_repair.htm";
			my $commandline = qq{start "$url" "$url"}; 	
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub cptOfficeRecoveryClicked {
			
			my $url = "http://www.godskingsandheroes.info/software/#5";
			my $commandline = qq{start "$url" "$url"}; 	
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub corruptOpenOfficeRecoveryClicked {
			
			my $url = "http://www.godskingsandheroes.info/software/#8";
			my $commandline = qq{start "$url" "$url"}; 	
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub recoveryForWordClicked {
			
			my $url = "http://www.officerecovery.com/word/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recovery for Excel's Web page with
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub recoveronixWordServiceClicked{
			
			my $url = "https://online.officerecovery.com/word/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recoveronix' online recovery with 
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub recoveryForExcelClicked {
			
			my $url = "http://www.officerecovery.com/excel/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recovery for Excel's Web page with
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		
		sub recoveryForPowerPointClicked {
			
			my $url = "http://www.officerecovery.com/powerpoint/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recovery for Excel's Web page with
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub recoveryForOpenOfficeWriterClicked {
			
			my $url = "http://www.officerecovery.com/writer/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recovery for Excel's Web page with
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub recoveryForOpenOfficeCalcClicked {
			
			my $url = "http://www.officerecovery.com/calc/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recovery for Excel's Web page with
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub recoveronixExcelServiceClicked{
			
			my $url = "https://online.officerecovery.com/excel/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recoveronix' online recovery with 
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub recoveronixPowerPointServiceClicked{
			
			my $url = "https://online.officerecovery.com/powerpoint/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recoveronix' online recovery with 
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}	
		
		sub recoveronixOpenOfficeWriterServiceClicked{
			
			my $url = "https://online.officerecovery.com/writer/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recoveronix' online recovery with 
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub recoveronixOpenOfficeCalcServiceClicked{
			
			my $url = "https://online.officerecovery.com/calc/?134994";
			my $commandline = qq{start "$url" "$url"}; 	# Launches Recoveronix' online recovery with 
			# my affiliate ID. Be sure to get your own.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub sendToMeClicked {
			
			my $url = "http://saveofficedata.com/contact.htm";
			my $commandline = qq{start "$url" "$url"}; # Launches contact form.
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}	
		
		sub wordFixClicked {
			
			my $url = "http://www.cimaware.com/info/info.php?id=622&path=main/products/wordfix.php";
			my $commandline = qq{start "$url" "$url"}; 	# Launches ExcelFix product page with my affiliate ID 622 :-). 
			# Change this to your affiliate ID with Cimaware...
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub excelFixClicked {
			
			my $url = "http://www.cimaware.com/info/info.php?id=622&path=main/products/excelfix.php";
			my $commandline = qq{start "$url" "$url"}; 	# Launches ExcelFix product page with my affiliate ID 622 :-). 
			# Change this to your affiliate ID with Cimaware...
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub corruptDocxSalvagerClicked {
			
			my $url = "http://www.godskingsandheroes.info/software/#4";
			my $commandline = qq{start "$url" "$url"}; 	# Launches ExcelFix product page with my affiliate ID 622 :-). 
			# Change this to your affiliate ID with Cimaware...
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub excelStepsClicked {
			
			my $url = "http://www.s2services.com/excel.htm";
			my $commandline = qq{start "$url" "$url"}; 	
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub openOfficeStepsClicked {
			
			my $url = "http://www.s2services.com/open_office.htm";
			my $commandline = qq{start "$url" "$url"}; 	
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub excelRecoveryClicked {
			
			my $url = "http://www.godskingsandheroes.info/software/#13";
			my $commandline = qq{start "$url" "$url"}; 	# Launches ExcelFix product page with my affiliate ID 622 :-). 
			# Change this to your affiliate ID with Cimaware...
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}
		
		sub corruptXlsxSalvagerClicked {
			
			my $url = "http://www.godskingsandheroes.info/software/#10";
			my $commandline = qq{start "$url" "$url"}; 	# Launches ExcelFix product page with my affiliate ID 622 :-). 
			# Change this to your affiliate ID with Cimaware...
			system($commandline) == 0
			or die qq{Couldn't launch '$commandline': $!/$?};
			
		}																																								