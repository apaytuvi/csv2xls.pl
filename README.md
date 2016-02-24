# csv2xlsx.pl
This script combine multiple `.csv` or `.tsv` files into different sheets in a `.xlsx` file.

Usage:

* `csv2xlsx.pl --in infile1.csv,infile2.csv,infile3.csv --out outfile.xlsx --worksheets infile1,infile2,infile3 --FS '\t' --header col1,col2,...,colX`

`--in` REQUIRED. Input files.

`--out` REQUIRED. `.xlsx` output file.

`--wordsheets` contains the name of each sheet separated by commas.

`--header` contains the name of each column for the header.

`--FS` is the field separator in the text files (for `.csv` files, it should be `','`; for `.tsv` (DEFAULT), it should be `'\t'`).
