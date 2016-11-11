# tancho
This is a D-Tools Price List Pre-Processor.

Its task is to take raw price-list .xls(x) files and convert them to an importable .csv

## Usage

Make csv files for parsing (depends on a functioning copy of MS Excel):
`XlsImport.vbs C:\Full\path\to\myFile.xlsx`

Create importable version of csv file (depends on Python 3):
```python
with PriceList("Name of Manufacturer", "myfile.csv") as foo:
	foo.parse()
	foo.write()
```