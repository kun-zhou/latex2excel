# latex2excel
This script takes a LaTeX file containing multiple tabular environments as input and outputs an Excel workbook version of all tables in which each worksheet contains one table. The package is currently compatible with `booktabs` package and `\multicolumn` commands, in addition to traditional commands like `\hline`. Try it out!

## Dependencies
`python3` is needed with `openpyxl` and `click` modules installed through pip. On mac, one can install all dependencies with
`brew install python3; pip3 install openpyxl; pip3 intall click;`

## Usage
*EXAMPLE*:
   `python3 latex2excel input_file [output_file]`
 
## Caveats:
- It does not support nested tables.
- Does not support nested commands in general because I used non-greedy regex.
