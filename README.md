# csv2xlsx.pl
This script combine multiple `.csv` or `.tsv` files into different sheets in a `.xlsx` file.

Usage:

* `csv2xlsx.pl --in infile1.csv,infile2.csv,infile3.csv --out outfile.xlsx --worksheets infile1,infile2,infile3 --FS '\t' --header col1,col2,...,colX`

`--in` input files (REQUIRED).

`--out` output file (REQUIRED).

`--wordsheets` sheet names.

`--header` column names for the header.

`--FS` field separator in the input files (for `.csv` files, it should be `','`; for `.tsv` (DEFAULT), it should be `'\t'`).
