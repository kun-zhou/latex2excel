# latex2excel
This script consumes a LaTeX source code file containing multiple tabular environments and generates an Excel workbook where each table is stored in an individual worksheet. In addition to traditional commands like `\hline` and `\multicolumn`, many commands from the `booktabs` package are also supported. Try it out!

## Dependencies
`python3` is used with `openpyxl` and `click` modules installed through pip. On macOS, one can install all dependencies with
`brew install python3; pip3 install openpyxl; pip3 intall click;`

## Usage
**EXAMPLE**:
   `python3 latex2excel input_file [output_file]`
 
## Caveats:
- It does not support nested tables.
- Does not support nested commands in general because this script uses simple regex to identify tabular environment and various commands.