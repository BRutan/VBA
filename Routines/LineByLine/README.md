------------------------------------------------------------------------------------------------------------------------------------------------
## Instructions to use the LineByLine Macro: 
------------------------------------------------------------------------------------------------------------------------------------------------
1. From the main worksheet click File -> Options -> Customize Ribbon -> Main Tabs 
then click the "Developer" checkbox. Click "Ok".
2. Save the Workbook as a "Excel Macro-Enabled Workbook" (.xlsm).
3. Sort by CPT in either ascending/descending order. 
4. Press Alt - F11 to open the VBA workspace.
5. Go to "File" -> "Import File" and select LineByLineForm.frm, then click "Open".
6. Go to "File" -> "Import File" and select LineByLine.bas, then click "Open".
7. Go to "File" -> "Import File" and select IntVector.cls, then click "Open".
8. Go to the Developer tab, click Insert (the toolbox picture) and select a Button.
9. Draw the button anywhere on the worksheet. Once drawn, you'll be asked to associate the button with a macro. Choose "YourWorkbookName.xlsb!Start_LineByLineForm".
10. Once you click the button the LineByLine userform will pop up. Enter the appropriate ranges into the input boxes then press the "Run" button.
