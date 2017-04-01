# latex2excel
 This script takes a latex file containing multiple tabular environments as input and outputs an Excel workbook version of all tables in which each worksheet contains one table. The package is currently compatible with `booktabs` package and `\multicolumn` commands, in addition to traditional commands like `\hline`. Try it out!
 
 
Catches:
1. It does not support nested tables.
2. Does not support nested commands in general because I used non-greedy regex instead of properly tokenize the latex file.
