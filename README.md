# editexcelpwm README

Want to edit Power Query M code of your xlsx file in VSCode and use Excel as interpriter? - Use my plugin. 
![Image of demo](images/demo.gif)

## Features

* Export all M queries from xlsx/xlsm file to *.m file
* Import queries from *.m file to xlsx/xlsm
* Edit M code in VSCode and run queries in Excel immediately 
* Create new queries and upload them to Excel



## Requirements

* VSCode ^1.33.0
* Windows
* MS Excel ^2016 - it uses AxtiveXObject to open xlsx and extract data

## Install to Visual Studio Code

1) Download [a relative link](editexcelpqm-1.0.0.vsix)
2) Go to download folder
3) Run in console **code --install-extension /path/to/vsix**

## Release Notes

### 1.0.0

Initial release. Adds the following commands to the context menue when you click on Excel file
* Export xlsx/xlsm -> *.m
* Import *.m -> xlsx/xlsm (without closing Excel and saving xlsx file)
* Save and close xlsx file
