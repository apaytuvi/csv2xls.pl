#!/usr/bin/perl
#Andreu Paytuví

use strict;
use warnings;
use Getopt::Long;
use Excel::Writer::XLSX;

my ($in, $out, $worksheets, $header, $help);
my $fieldseparator = "\t";

GetOptions ("in=s" => \$in,
            "out=s" => \$out,
            "worksheets:s" => \$worksheets,
            "header:s" => \$header,
            "FS:s" => \$fieldseparator,
            "help" => \$help
            )
or die("Error in command line arguments.\nSee 'script-pl --help' 'script.pl -h' for further information.\n\n");

sub usage {
  print "csv2xls --in infile1.csv,infile2.csv,infile3.csv [REQUIRED] --out outfile.xlsx [REQUIRED] --worksheets infile1,infile2,infile3 --header col1,col2,...,colX --FS '\\t'\n\n";
  die;
}

if ($help || !defined $in || !defined $out) {
  usage();
}

my @files = split ",", $in;
my @header_names;
my @sheets;

if ($worksheets) {
  @sheets = split ",", $worksheets;
  if ($#files != $#sheets) {
    usage();
  }
}

my $add = 0;
if ($header) {
  @header_names = split ",", $header;
  $add = 1;
}

my $workbook = Excel::Writer::XLSX->new($out);
my $format = $workbook->add_format();
$format->set_bold(1); #header will be bold
$format->set_bg_color("#E2E0DD"); 

for my $n (0 .. $#files) {
  open (FILE, $files[$n]);
  my %string_len; #here we will store string lengths
  my $worksheet = $workbook->add_worksheet($sheets[$n]);
  print "File .. ".$files[$n]." .. to sheet .. ".$sheets[$n]."\n";
  my $row = 0;
  while (my $line = <FILE>) {
    chop $line;
    my @Fld = split $fieldseparator, $line;
    my $col = 0;
    for my $token (@Fld) {
      if ($row == 0) {
        $string_len{$col} = [];
        if ($header) {
          $worksheet->write($row, $col, $header_names[$col], $format);
          push @{$string_len{$col}}, length($header_names[$col]);
        }
      }
      $worksheet->write($row+$add, $col, $token);
      push @{$string_len{$col}}, length($token);
      $col++;
    }
    $row++;
  }
  for my $key (keys %string_len) { #controlling col width
    my $col_width;
    my @sorted = sort {$b <=> $a} @{$string_len{$key}};
    if ($sorted[0] > 50) {
      $col_width = 50;
    } else {
      $col_width = $sorted[0];
    }
    $worksheet->set_column($key, $key, $col_width);
  }
}
