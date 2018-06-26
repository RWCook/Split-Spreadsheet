# Split-Spreadsheet

Excel macro to split a spreadsheet into different groups and to output these groups to separate files, or to to different worksheets within one file. This might, for example, be useful in the scenario where you need to split a spreadsheet in order to send different files to different people.

The macro loads a form which prompts for 
* a data source file (i.e. the spreadsheet to split)
* an output directory to write the results to
* the name of the sheet that contains the data
* whether to write the data to new workbooks or to worksheets within a workbook
* which column to split on (by column number)

Additionally, there are options to apply basic formatting (which formats a heading, autofits the cell width, and adds borders) and to stop Excel warning before it overwrites files.

The macro makes a copy of the data source file, in case there are any problems, and creates named ranges for each group of data. These ranges are then copied in to the new files or worksheets. 

The macro works on the assumption that the data has been sorted by the split column; any further levels of sorting are maintained in the output.

