# Excel Stacker

Excel Stacker allows you to merge data from multiple Excel files into one file. It locates all the .xls and .xlsx files
in a selected folder and stacks rows from each file ontop of each other.

## Usage

Available options when using program:

  - `Choose Folder`: Allows you to select the folder where the .xls and .xlsx files are
  - `Skip Rows`: Allows you to skip the first `x` number of rows from each speadsheet, ie. remove headings
  - `Output File Name`: Name of the megered spreadsheet. Saved in the same directory as the app

## Converting Excel Stacker to .exe

  Converting .py application to .exe was done using Python library Pyinstaller, example in `dist` folder. The below command run in the terminal to create the .exe application
  
  ```
  pyinstaller excel_stacker.py --onefile --windowed --icon=img/icon.ico
  ```
  
