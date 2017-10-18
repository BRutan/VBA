------------------------------------------------------------------------------------------------------------------------------------------------
## RangeDifference Macro Description: 
------------------------------------------------------------------------------------------------------------------------------------------------
Macro calculates the "difference" (set difference) between two ranges (A and B). The difference is all strings in set A that are NOT in set B (set notation: A \ B). 

------------------------------------------------------------------------------------------------------------------------------------------------
## RangeDifference Macro Import Instructions: 
------------------------------------------------------------------------------------------------------------------------------------------------
1. From the main worksheet click File -> Options -> Customize Ribbon -> Main Tabs 
then click the "Developer" checkbox. Click "Ok".
2. Save the Workbook as a "Excel Macro-Enabled Workbook" (.xlsm).
3. Press Alt - F11 to open the VBA workspace.
4. Go to "File" -> "Import File" and select RangeDifference.frm, then click "Open".
5. Go to "File" -> "Import File" and select RangeDifference.bas, then click "Open".
6. Go to "File" -> "Import File" and select RangeDifference.frx, then click "Open".
7. Go to the Developer tab, click Insert (the toolbox picture) and select a Button.
8. Draw the button anywhere on the worksheet. Once drawn, you'll be asked to associate the button with a macro. Choose "YourWorkbookName.xlsb!Start_RangeDifferenceForm".
9. Once you click the button the RangeDifference userform will pop up. Enter the appropriate ranges into the input boxes then press the "Run" button. See RangeDifference macro content in VBA workspace for further details. 
