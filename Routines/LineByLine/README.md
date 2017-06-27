------------------------------------------------------------------------------------------------------------------------------------------------
## Instructions to use the LineByLine Macro: 
------------------------------------------------------------------------------------------------------------------------------------------------
1. (If you haven't already:) From the main worksheet click File -> Options -> Customize Ribbon -> Main Tabs 
then click the "Developer" checkbox. Click "Ok".
2. Save the Workbook as a "Excel Macro-Enabled Workbook" (.xlsm).
3. Insert a "Group #" column to the right of the CPT column in the pricing file. Run a VLookup on CPT to get the Group # in the CPT-Group # rollup. (TBD)
4. Sort by CPT in either ascending/descending order. 
5. Press Alt - F11 to open the VBA workspace.
6. Go to "File" -> "Import File" and select LineByLineForm.frm, then click "Open".
7. Go to "File" -> "Import File" and select LineByLine.bas, then click "Open".
8. Go to "File" -> "Import File" and select IntVector.cls, then click "Open".
9. Go to the Developer tab, click Insert (the toolbox picture) and select a Button.
10. Draw the button anywhere on the worksheet. Once drawn, you'll be asked to associate the button with a macro. Choose "YourWorkbookName.xlsb!Start_LineByLineForm".
11. Once you click the button the LineByLine userform will pop up. Enter the appropriate ranges into the input boxes then press the "Run" button.
