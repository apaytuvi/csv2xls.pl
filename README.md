# csv2xls.pl
These scripts combine multiple `.csv` or `.tsv` files into different sheets of a `.xls` or `.xlsx` file.

Usage:

`csv2xls.pl --in infile1.csv,infile2.csv,infile3.csv --out outfile.xls --worksheets infile1,infile2,infile3 --FS '\t'`
`csv2xlsx.pl --in infile1.csv,infile2.csv,infile3.csv --out outfile.xlsx --worksheets infile1,infile2,infile3 --FS '\t'`

`--wordsheets` contains the name of each sheet separated by commas.

`--FS` is the field separator in the text files (for `.csv` files, it should be `','`; for `.tsv`, it should be `'\t'`).
