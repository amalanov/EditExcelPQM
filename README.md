# EditExcelPQM - edit M code of your xlsx in VSCode
Want to edit Power Query M code of your xlsx file in VSCode and use Excel as interpriter? - Here is the plugin. 

## Features
* Export all M queries from xlsx/xlsm file to *.m file
* Import queries from *.m file to xlsx/xlsm
* Edit M code in VSCode and run queries in Excel immediately 
* Create new queries and upload them to Excel
* Delete queries from VSCode

## Demo
![Image of demo](images/demo.gif)

## Install to Visual Studio Code
From [VSCode extensions market](https://marketplace.visualstudio.com/items?itemName=AMalanov.editexcelpqm) or manually:
1) Download [vsix file](editexcelpqm-1.0.1.vsix) from this repo
2) Go to download folder
3) Run in console **code --install-extension /path/to/vsix**

## Known issues
* Unable to fully close Excel - window is closed, but it remains in process manager
* If your Excel shows a message on startup, plugin is unable to access queries before you close the popup
* On some systems the plugin opens Excel in background mode and I'm not able to do it visible.

## Requirements
* VSCode ^1.33.0
* Windows
* MS Excel ^2016 - it uses AxtiveXObject to open xlsx and extract data