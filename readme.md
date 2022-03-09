** coded by a spacecowboy **

Notes:

+ The script uses the Drive service. The user must install it manually
+ The script will override the active spreadsheet's content in the output sheet, and it will only write into the three first columns.
+ The script expects the active spreadsheet to be blank.


+ The script starts when the output/results sheet is opened, user can then choose to stop the script from doing its thing, or let it run!
+ The script uses the default folder to read excel files from if user clicks exit button in the "where have you put the excel files" dialog
+ The script creates a custom menu
+ The script uses a tmp folder in the root of the drive, if none is found, the script will create one!
+ If multiple tmp folders exist, the script will use the first one!
----------
+ The script creates temporary google sheet for every excel file,it deletes them once data extraction is done.
+ The user will be alerted when cleanup function is triggered (this happens in case the script fails to clear the converted excel sheets)
----------
+ "Script.gs" file contains all the project, so you can copy paste into google script IDE directly!
