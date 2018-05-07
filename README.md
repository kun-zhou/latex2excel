# latex2excel
This script consumes a LaTeX source code file containing multiple tabular environments and generates an Excel workbook where each table is stored in an individual worksheet. In addition to traditional commands like `\hline` and `\multicolumn`, many commands from the `booktabs` package are also supported. I wrote this script during my tenure as a research assistant.

Try it out!

## Install
Make sure you have `pip` installed. Then run `pip install latex2excel` or `pip3 install latex2excel`.

### Dummies' Guide for macOS Users
If on macOS, run the following:
`/usr/bin/ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)"`
`brew install python3`
`pip3 install latex2excel`
in a terminal window. this sets up Homebrew and installs python3 using Homebrew.

## Usage
   - `latex2excel recent_table.tex` will generate an Excel file with the name `recent_table.xlsx` at same directory as `recent_table.tex`
   - `latex2excel recent_table.tex table_03` will generate an Excel file with the name `table_03.xlsx` in the current working directory (use `pwd` to get the current working directory)
 
## Caveats:
- It does not support nested tables.
- Does not support nested commands in general because this script uses simple regex to identify tabular environment and various commands.
