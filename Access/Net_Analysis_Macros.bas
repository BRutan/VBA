Attribute VB_Name = "Net_Analysis_Macros"
Option Compare Database
Option Explicit
Sub ImportNATables()
'' IN PROGRESS:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Convert spreadsheets from a passed NA Rate Matrix Excel workbook into Net Analysis tables
'' automatically.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Todo:
'' 1. Convert workbook extension to file format object.
'' 2. Replace tables with same name already in current database.
'' 3. Prevent Excel workbook from being displayed when calling the TransferSpreadsheet method.
'' 4. Use CreateObject() to instantiate excel objects to ensure portability.

Dim startWSNum As Integer
Dim errorString As String: errorString = "Error:"
Dim hasError As Boolean
Dim excelApp As Object
Dim inputWB As Object
Dim workbookPathString As String: workbookPathString = InputBox("Enter (path to NA breakout workbook, start page num): ")


Call Access_Macro_Utilities.OptimizeCodeSettings(True)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Check inputs:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If StrComp(Trim(workbookPathString), "") = 0 Then
    errorString = errorString & vbCr & vbTab & "Please enter the input to the prompt."
    hasError = True
    GoTo Exit_With_Error
End If
If InStr(1, workbookPathString, ",") = 0 Then
    errorString = errorString & vbCr & vbTab & "Please enter a number following a comma following the file path."
    hasError = True
ElseIf IsNumeric(Mid(workbookPathString, InStr(1, workbookPathString, ",") + 1, Len(workbookPathString))) = False Then
    errorString = errorString & vbCr & vbTab & "Starting page " & Chr(34) & Mid(workbookPathString, InStr(1, workbookPathString, ",")) & Chr(34) & " is invalid."
    hasError = True
Else
    startWSNum = CInt(Mid(workbookPathString, InStr(1, workbookPathString, ",") + 1, Len(workbookPathString)))
    workbookPathString = Mid(workbookPathString, 1, InStr(1, workbookPathString, ",") - 1)
End If

' Open Excel application instance:
' Check if excel already running. If not then spawn new Excel instance.
'Set excelApp = GetObject(, "Excel.Application")
 
If excelApp Is Nothing Then
    ' Spawn new Excel instance.
    Set excelApp = CreateObject("Excel.Application")
End If

excelApp.Visible = False
excelApp.DisplayAlerts = False

If Err.Number <> 0 Then
    errorString = errorString & vbCr & vbTab & "Excel failed to load."
    hasError = True
    Err.Clear
Else
' Open target workbook:
'''''''' TODO: investigate why code returns to this part.

On Error GoTo Cont
    Set inputWB = excelApp.Workbooks.Open(workbookPathString)
Cont:
    If Err.Number <> 0 Then
        errorString = errorString & vbCr & vbTab & "Workbook path " & Chr(34) & workbookPathString & Chr(34) & " is invalid."
        hasError = True
        Err.Clear
    End If
    ' If starting worksheet number is beyond all worksheets then add error:
    If startWSNum > inputWB.Worksheets.Count Then
        errorString = errorString & vbCr & vbTab & "Worksheet number " & CStr(startWSNum) & " is out of bounds (" & CStr(inputWB.Worksheets.Count) & " worksheets total). "
        hasError = True
    End If
End If

' Exit macro if there were input errors:
If hasError = True Then
    GoTo Exit_With_Error
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Main Routine:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim thisDB As DAO.Database: Set thisDB = Application.CurrentDb
Dim inputWBFileType As AcSpreadSheetType

Dim currWSNum As Integer
For currWSNum = startWSNum To inputWB.Worksheets.Count
    ' If worksheetname (as table) is already in database then delete it:
    DoCmd.DeleteObject acTable, thisDB.TableDefs(inputWB.Worksheets(currWSNum).name)
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, inputWB.Worksheets(currWSNum).name, workbookPathString, True, inputWB.Worksheets(currWSNum).name & "$"
Next currWSNum

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Cleanup, saving:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GoTo Exit_Sub

Exit_Sub:
    excelApp.Quit
    Set excelApp = Nothing
    Set inputWB = Nothing
    Call Access_Macro_Utilities.OptimizeCodeSettings(False)
    Exit Sub
    
Exit_With_Error:
    excelApp.Quit
    Set excelApp = Nothing
    Set inputWB = Nothing
    Call Access_Macro_Utilities.OptimizeCodeSettings(False)
    If hasError = True Then
        MsgBox errorString
    Else
        MsgBox "Other macro error."
    End If
    Exit Sub

End Sub

Sub AppendNAPriceDiffFields()
'''''''''''''' IN PROGRESS:
'' Errors: appending multiple fields causes macro to fail.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Note: To work, must set the Net Analysis table name to "Net Analysis".

Dim activeDB As DAO.Database: Set activeDB = CurrentDb
' Skip to the ExitSub label if the 'Net Analysis' table does not exist, or some other error occurs:
On Error GoTo Exit_With_Error

' Turn off inefficient Access settings:
Call Access_Macro_Utilities.OptimizeCodeSettings(True)

Dim NATable As DAO.TableDef
If Access_Macro_Utilities.CheckObjectOpen("Net Analysis", 0) = True Then
    DoCmd.Close acTable, "Net Analysis", acSaveYes
End If
Set NATable = activeDB.TableDefs("Net Analysis")
Dim currTableDef2 As DAO.TableDef
Dim currField As DAO.field
Dim nameString As String
Dim needToAddLast3 As Boolean: needToAddLast3 = True

Dim namesArray() As String
Dim currArrayIndex As Integer: currArrayIndex = 0

' Generate the table update fields based upon tables from Rate Matrix:
For Each currTableDef2 In activeDB.TableDefs
    ' Get all characters to the right up dashes:
    If InStr(currTableDef2.name, "IP ") <> 0 Or InStr(currTableDef2.name, " IP") <> 0 Or _
        InStr(currTableDef2.name, "OP ") <> 0 Or InStr(currTableDef2.name, "OP ") <> 0 Then
        nameString = Trim(Mid(currTableDef2.name, InStr(currTableDef2.name, "-") + 1))
        If Mid(nameString, 1, 2) = "OP" And needToAddLast3 = True Then
            ' Add Final, Dummy Gross and Dummy Net columns if do not exist:
            ReDim Preserve namesArray(currArrayIndex + 2)
            namesArray(currArrayIndex) = "IP Final"
            namesArray(currArrayIndex + 1) = "IP Dummy Gross"
            namesArray(currArrayIndex + 2) = "IP Dummy Net"
            currArrayIndex = currArrayIndex + 2
            needToAddLast3 = False
        End If
        ReDim Preserve namesArray(currArrayIndex)
        namesArray(currArrayIndex) = nameString
        currArrayIndex = currArrayIndex + 1
    End If
Next currTableDef2

' Add the OP Final, OP Dummy Gross, and OP Dummy Net columns:
ReDim Preserve namesArray(currArrayIndex + 2)
namesArray(currArrayIndex) = "OP Final"
namesArray(currArrayIndex + 1) = "OP Dummy Gross"
namesArray(currArrayIndex + 2) = "OP Dummy Net"
currArrayIndex = currArrayIndex + 2
' Add the Price Diff, IP/OP Gross, Total Gross, IP/OP Net and Total Net Columns
ReDim Preserve namesArray(currArrayIndex + 6)
namesArray(currArrayIndex) = "Price Diff"
namesArray(currArrayIndex + 1) = "IP Gross"
namesArray(currArrayIndex + 2) = "OP Gross"
namesArray(currArrayIndex + 3) = "Total Gross"
namesArray(currArrayIndex + 4) = "IP Net"
namesArray(currArrayIndex + 5) = "OP Net"
namesArray(currArrayIndex + 6) = "Total Net"
currArrayIndex = currArrayIndex + 6

' Delete columns if already present in table:
For currArrayIndex = 0 To UBound(namesArray)
    If Access_Macro_Utilities.CheckFieldExists(NATable.name, namesArray(currArrayIndex)) = True Then
        NATable.Fields.Delete namesArray(currArrayIndex)
        NATable.Fields.Refresh
    End If
Next currArrayIndex

' Append columns:
''''' TODO: investigate why failing on second iteration.
For currArrayIndex = 0 To UBound(namesArray)
    NATable.CreateField namesArray(currArrayIndex), dbDouble
    NATable.Fields.Refresh
Next currArrayIndex

' Nothing out all the referenced objects and exit:
Set activeDB = Nothing
Set NATable = Nothing
Set currTableDef2 = Nothing
Exit Sub

' Subroutine goes here if error occurs.
Exit_With_Error:
    Set activeDB = Nothing
    Set NATable = Nothing
    Set currTableDef2 = Nothing
    MsgBox Err.Description
    Call Access_Macro_Utilities.OptimizeCodeSettings(False)
    Exit Sub

End Sub

Sub ImportPrevYearNATableStructure()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Import JUST the 'Net Analysis' table from the previous year's Net Analysis model.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim databasePath As String: databasePath = InputBox("Enter the path to the previous year's Net Analysis database:")
' Turn off inefficient Access settings:

Dim errorString As String: errorString = "Error:"
Dim hasError As Boolean

If StrComp(Trim(databasePath), "") = 0 Then
    errorString = errorString & vbCr & vbTab & "Please enter the path to the previous year's database."
    hasError = True
    GoTo Exit_With_Error
End If

Call Access_Macro_Utilities.OptimizeCodeSettings(True)

On Error GoTo Exit_With_Error

' Determine if previous year's database has a Net Analysis table:
Dim thisDB As DAO.Database: Set thisDB = Application.CurrentDb
Dim prevYrDatabase As DAO.Database: Set prevYrDatabase = DBEngine.Workspaces(0).OpenDatabase(databasePath, ReadOnly:=False)
Dim naTableName As String

Dim currTable As DAO.TableDef
Dim hasNATable As Boolean
For Each currTable In prevYrDatabase.TableDefs
    If InStr(1, Trim(currTable.name), "Net Analysis") <> 0 Then
        hasNATable = True
        naTableName = currTable.name
        Exit For
    End If
Next currTable

If hasNATable = False Then
    errorString = errorString & vbCr & vbTab & "Database does not have a table titled Net Analysis."
    hasError = True
End If

If hasError = True Then
    GoTo Exit_With_Error
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Main Routine:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Delete the Net Analysis table in current database if exists.
If Not (thisDB.TableDefs("Net Analysis") Is Nothing) Then
    thisDB.Execute "DROP TABLE [Net Analysis]"
End If
' Import the Net Analysis table (structure only) from the passed database:
DoCmd.TransferDatabase acImport, DataBaseType:="Microsoft Access", DatabaseName:=databasePath, ObjectType:=acTable, Source:=naTableName, Destination:="Net Analysis", StructureOnly:=True

Set thisDB = Nothing
Set prevYrDatabase = Nothing

Call Access_Macro_Utilities.OptimizeCodeSettings(False)

Exit Sub

Exit_With_Error:
    Set thisDB = Nothing
    Set prevYrDatabase = Nothing
    Call Access_Macro_Utilities.OptimizeCodeSettings(False)
    ' Display error message:
    If hasError = True Then
        MsgBox errorString
    Else
        MsgBox Err.Description
    End If
    
End Sub

Sub ImportNALastYrQueries()
'''''''''''' IN PROGRESS:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Import all queries from last year's NA database into current NA model.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim databasePath As String: databasePath = InputBox("Enter the path to the previous year's Net Analysis database:")
Dim errorString As String: errorString = "Error:"
Dim thisDB As DAO.Database
Dim inputDB As DAO.Database
Dim importQueries() As String
Dim currQuery As QueryDef
Dim hasError As Boolean
Dim newQuery As QueryDef

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Input checking:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If StrComp(Trim(databasePath), "") = 0 Then
    errorString = errorString & vbCr & vbTab & "Please enter the path to the previous year's database."
    hasError = True
    GoTo Exit_With_Error
End If

If DBEngine.OpenDatabase(databasePath, ReadOnly:=False) Is Nothing Then
    errorString = errorString & vbCr & vbTab & "Database path " & Chr(34) & databasePath & Chr(34) & " is invalid."
    hasError = True
End If

If hasError = True Then
    GoTo Exit_With_Error
End If
' Turn off inefficient Access settings:
'Call Access_Macro_Utilities.OptimizeCodeSettings(True)

On Error GoTo Exit_With_Error
Set thisDB = Application.CurrentDb
Set inputDB = DBEngine.OpenDatabase(databasePath, ReadOnly:=False)
Dim totalQueries As Integer: totalQueries = inputDB.QueryDefs.Count
Dim i As Integer
ReDim importQueries(totalQueries - 1, 1)

' Make copies of all the inputDB queries:
For i = 0 To totalQueries - 1
    importQueries(i, 0) = inputDB.QueryDefs(i).name
    importQueries(i, 1) = inputDB.QueryDefs(i).SQL
Next i

inputDB.Close

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Main Method: Import queries from specified database.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Import all queries from the target database:
For i = 0 To totalQueries - 1
    ' Delete query if exists in the active database:
    If Access_Macro_Utilities.CheckQueryExists(importQueries(i, 0)) = True Then
        thisDB.QueryDefs(importQueries(i, 0)).SQL = importQueries(i, 1)
    Else
    ' Append query (must convert to SQL to append):
        thisDB.CreateQueryDef importQueries(i, 0), importQueries(i, 1)
        thisDB.QueryDefs.Refresh
    End If
Next i

' Nothing out objects and exit:
Set thisDB = Nothing
Set inputDB = Nothing
Set currQuery = Nothing

Call Access_Macro_Utilities.OptimizeCodeSettings(False)

Exit Sub

Exit_With_Error:
  If Not (inputDB Is Nothing) Then
    inputDB.Close
  End If
  Set thisDB = Nothing
  Set inputDB = Nothing
  Set currQuery = Nothing

  Call Access_Macro_Utilities.OptimizeCodeSettings(False)
  If hasError = True Then
    MsgBox errorString
  Else
    MsgBox Err.Description
  End If

End Sub
