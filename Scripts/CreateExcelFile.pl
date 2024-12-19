#!/usr/bin/perl  
use strict;
use warnings;

use Excel::Writer::XLSX; 

my $Excelbook = Excel::Writer::XLSX->new( 'C:\Users\jeffl\OneDrive\Documents\GitHub\excel-automation\Spreadsheets\Excel Perl Test.xlsx' ); 
my $Excelsheet = $Excelbook->add_worksheet(); 

$Excelsheet->write( "A1", "This is Cell A1" ); 

$Excelbook->close; 