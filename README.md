# csv2xls.pl
This script combine multiple `.csv` or `.tsv` files into different sheets of a `.xls` file

Usage:

`csv2xls.pl --in infile1.csv,infile2.csv,infile3.csv --out outfile.xls --worksheets infile1,infile2,infile3 --FS '\t'`

`--wordsheets` are the name of each sheet separated by commas.

`--FS` is the field separator in the text files (for `.csv` files, it should be ','; for `.tsv`, it should be '\t').
