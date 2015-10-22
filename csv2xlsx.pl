#!/usr/bin/perl
#Andreu Paytuví

use strict;
use warnings;
use Getopt::Long;
use Excel::Writer::XLSX;

my ($in, $out, $worksheets, $help);
my $fieldseparator = "\t";

GetOptions ("in=s" => \$in,
            "out=s" => \$out,
            "worksheets:s" => \$worksheets,
            "FS:s" => \$fieldseparator,
            "help" => \$help
            )
or die("Error in command line arguments.\nSee 'script-pl --help' 'script.pl -h' for further information.\n\n");

sub usage {
  print "csv2xls --in infile1.csv,infile2.csv,infile3.csv --out outfile.xlsx --worksheets infile1,infile2,infile3 --FS '\\t'\n\n";
  die;
}

if ($help || !defined $in || !defined $out) {
  usage();
}

my @files = split ",", $in;
my @sheets;

if ($worksheets) {
  @sheets = split ",", $worksheets;
  if ($#files != $#sheets) {
    usage();
  }
}

my $workbook = Excel::Writer::XLSX->new($out);

for my $n (0 .. $#files) {
  open (FILE, $files[$n]);
  my $worksheet = $workbook->add_worksheet($sheets[$n]);
  print "File .. ".$files[$n]." .. to sheet .. ".$sheets[$n]."\n";
  my $row = 0;
  while (my $line = <FILE>) {
    chop $line;
    my @Fld = split $fieldseparator, $line;
    my $col = 0;
    for my $token (@Fld) {
      $worksheet->write($row, $col, $token);
      $col++;
    }
    $row++;
  }
}

