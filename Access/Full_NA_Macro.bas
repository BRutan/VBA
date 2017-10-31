Attribute VB_Name = "FullNAMacro"
Option Compare Database
Option Explicit
''''''''''''''' TODO:
'' 1. Add ability to skip step.

Sub MainMethod()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Run this subroutine to perform all of the Net Analysis preparation:
'' 1. Import all tables from THIS YEAR's Net Analysis broken-out rate matrix.
'' 2. Import structure of PREVIOUS YEAR's net analysis table.
'' 3. Import all PREVIOUS YEAR's queries.
'' Note: The queries will need to be updated to correspond to this year's tables and the Net Analysis table may need
'' to have columns added.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' Programmatically add reference to Excel objects:
Dim ref As Reference
Dim hasRef As Boolean
Dim currRef As Reference

' Determine if already present:
For Each currRef In Access.References
    If StrComp(currRef.name, "Excel") = 0 Then
        hasRef = True
        Exit For
    End If
Next currRef

' Add the reference to Excel if not already present (note: assumes using Office 2016):
If hasRef = False Then
    Access.References.AddFromFile "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
End If


'' Determine if user wants to clear out the current database:
Dim clearDBResult As Integer: clearDBResult = MsgBox("Clear current database of all queries and tables?", vbOKCancel)
Dim currTable As DAO.TableDef
Dim currQuery As DAO.QueryDef
Dim tempSTR As String

If clearDBResult = 1 Then
    '' Delete all tables:
    For Each currTable In CurrentDb.TableDefs
    '' Determine if table not a system table:
        If InStr(1, currTable.name, "Msys") = 0 Then
         '' Close table if open:
            If Access_Macro_Utilities.CheckObjectOpen(currTable.name, 0) = True Then
                DoCmd.Close acTable, currTable.name, acSaveNo
            End If
            DoCmd.DeleteObject acTable, currTable.name
        End If
    Next currTable
    '' Delete all queries:
    For Each currQuery In CurrentDb.QueryDefs
        '' Close query if open:
        If Access_Macro_Utilities.CheckObjectOpen(currQuery.name, 1) = True Then
            DoCmd.Close acQuery, currQuery.name, acSaveNo
        End If
        DoCmd.DeleteObject acQuery, currQuery.name
    Next currQuery
    Set currTable = Nothing
    Set currQuery = Nothing
End If

'' Programmatically add reference to the Microsoft Excel Object DLL:

Dim dbPath As String: dbPath = InputBox("Enter Last Year's Net Analysis database path:")
Dim excelInputString As String: excelInputString = InputBox("Enter path for Net Analysis Excel Rate Matrix and starting worksheet number in form" & vbCr & " workbookPath, startNum ")
Dim errorString As String: errorString = "Net Analysis Macro Error: "
Dim newError As String
Dim hasError As Boolean

If InStr(1, excelInputString, ",") = 0 Then
    errorString = vbCr & "Please enter Net Analysis Excel Rate Matrix and starting worksheet number in form" & vbCr & " workbookPath, startNum "
    hasError = True
    GoTo Exit_With_Error
End If

Call Access_Macro_Utilities.OptimizeCodeSettings(True)

'' Import the Net Analysis tables from the passed Excel spreadsheet:
newError = ImportNATables(excelInputString)
If StrComp(Trim(newError), "") <> 0 Then
    errorString = errorString & vbCr & newError
    hasError = True
End If
hasError = False

'' Import the Net Analysis table from the previous year's database:
newError = ImportPrevYearNATableStructure(dbPath)
If StrComp(Trim(newError), "") <> 0 Then
    errorString = errorString & vbCr & newError
    hasError = True
End If

'' Import the Net Analysis queries from the previous year's database:
newError = ImportNALastYrQueries(dbPath)
If StrComp(Trim(newError), "") <> 0 Then
    errorString = errorString & vbCr & newError
    hasError = True
End If
hasError = False
If hasError = True Then
    GoTo Exit_With_Error
End If

'' Exit subroutine:

Call Access_Macro_Utilities.OptimizeCodeSettings(False)

MsgBox "Net Analysis Excel Tables, Queries and Net Analysis Access Table have been imported." & vbCr & "Note: Some tables will have imported additional empty rows, just delete them out."

Exit Sub

Exit_With_Error:
    Call Access_Macro_Utilities.OptimizeCodeSettings(False)
    If hasError = True Then
        MsgBox errorString
    Else
        MsgBox "Error: " & Err.Description
    End If
End Sub

Function ImportNATables(excelPathString As String) As String
'' IN PROGRESS:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Convert spreadsheets from a passed NA Rate Matrix Excel workbook into Net Analysis tables
'' automatically.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Todo:
'' 1. Convert workbook extension to file format object.

Dim errorString As String: errorString = "Import Net Analysis Excel Tables Error:"
Dim startWSNum As Integer
Dim hasError As Boolean
Dim excelApp As Excel.Application
Dim inputWB As Excel.Workbook
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Check inputs:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If StrComp(Trim(excelPathString), "") = 0 Then
    errorString = errorString & vbCr & "Please enter the input to the prompt."
    hasError = True
    GoTo Exit_With_Error
End If
If InStr(1, excelPathString, ",") = 0 Then
    errorString = errorString & vbCr & "Please enter a number following a comma following the file path."
    hasError = True
ElseIf IsNumeric(Mid(excelPathString, InStr(1, excelPathString, ",") + 1, Len(excelPathString))) = False Then
    errorString = errorString & vbCr & "Starting page " & Chr(34) & Mid(excelPathString, InStr(1, excelPathString, ",")) & Chr(34) & " is invalid."
    hasError = True
Else
    startWSNum = CInt(Mid(excelPathString, InStr(1, excelPathString, ",") + 1, Len(excelPathString)))
    excelPathString = Mid(excelPathString, 1, InStr(1, excelPathString, ",") - 1)
End If


' Open Excel application instance:
' Check if excel already running. If not then spawn new Excel instance.
'Set excelApp = GetObject(, "Excel.Application")
 
If excelApp Is Nothing Then
    ' Spawn new Excel instance.
    Set excelApp = New Excel.Application
    'Set excelApp = CreateObject("Excel.Application")
End If

excelApp.Visible = False
excelApp.DisplayAlerts = False

If Err.Number <> 0 Then
    errorString = errorString & vbCr & "Excel failed to load."
    hasError = True
    Err.Clear
Else
' Open target workbook:
On Error GoTo Cont
' Figure out why failing:
    Set inputWB = excelApp.Workbooks.Open(excelPathString)
Cont:
    If Err.Number <> 0 Then
        errorString = errorString & vbCr & "Workbook path " & Chr(34) & excelPathString & Chr(34) & " is invalid."
        hasError = True
        Err.Clear
    End If
    ' If starting worksheet number is beyond all worksheets then add error:
    If startWSNum > inputWB.Worksheets.Count Then
        errorString = errorString & vbCr & "Worksheet number " & CStr(startWSNum) & " is out of bounds (" & CStr(inputWB.Worksheets.Count) & " worksheets total). "
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
    If Access_Macro_Utilities.CheckObjectExists(inputWB.Worksheets(currWSNum).name, 0) = True Then
        DoCmd.DeleteObject acTable, thisDB.TableDefs(inputWB.Worksheets(currWSNum).name)
    End If
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, inputWB.Worksheets(currWSNum).name, excelPathString, True, inputWB.Worksheets(currWSNum).name & "$"
Next currWSNum

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Cleanup, saving:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GoTo Exit_Sub

Exit_Sub:
    excelApp.Quit
    Set excelApp = Nothing
    Set inputWB = Nothing
    Exit Function
    
Exit_With_Error:
    excelApp.Quit
    Set excelApp = Nothing
    Set inputWB = Nothing
    If hasError = True Then
        ImportNATables = errorString
    Else
        ImportNATables = errorString & vbCr & Err.Description
    End If
    Exit Function

End Function
Function ImportPrevYearNATableStructure(dbPath As String) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Import JUST the 'Net Analysis' table from the previous year's Net Analysis model.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Turn off inefficient Access settings:

Dim errorString As String: errorString = "Import Net Analysis Table Error:"
Dim hasError As Boolean

If StrComp(Trim(dbPath), "") = 0 Then
    errorString = errorString & vbCr & "Please enter the path to the previous year's database."
    hasError = True
    GoTo Exit_With_Error
End If

On Error GoTo Exit_With_Error

' Determine if previous year's database has a Net Analysis table:
Dim thisDB As DAO.Database: Set thisDB = Application.CurrentDb
Dim prevYrDatabase As DAO.Database: Set prevYrDatabase = DBEngine.Workspaces(0).OpenDatabase(dbPath, ReadOnly:=False)
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
    errorString = errorString & vbCr & "Database does not have a table titled Net Analysis."
    hasError = True
End If

If hasError = True Then
    GoTo Exit_With_Error
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Main Routine:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Delete the Net Analysis table in current database if exists.
If Access_Macro_Utilities.CheckObjectExists(naTableName, 0) = True Then
    thisDB.Execute "DROP TABLE [Net Analysis]"
End If
' Import the Net Analysis table (structure only) from the passed database:
DoCmd.TransferDatabase acImport, DataBaseType:="Microsoft Access", DatabaseName:=dbPath, ObjectType:=acTable, Source:=naTableName, Destination:="Net Analysis", StructureOnly:=True

Set thisDB = Nothing
Set prevYrDatabase = Nothing

ImportPrevYearNATableStructure = ""

Exit Function

Exit_With_Error:
    Set thisDB = Nothing
    Set prevYrDatabase = Nothing
    ' Display error message:
    If hasError = True Then
        ImportPrevYearNATableStructure = errorString
    Else
        ImportPrevYearNATableStructure = errorString & vbCr & Err.Description
    End If
    
End Function

Function ImportNALastYrQueries(dbPath As String) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Import all queries from last year's NA database into current NA model.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim errorString As String: errorString = "Import Net Analysis Queries Error:"
Dim thisDB As DAO.Database
Dim inputDB As DAO.Database
Dim importQueries() As String
Dim currQuery As QueryDef
Dim hasError As Boolean
Dim newQuery As QueryDef

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Input checking:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If StrComp(Trim(dbPath), "") = 0 Then
    errorString = errorString & vbCr & "Please enter the path to the previous year's database."
    hasError = True
    GoTo Exit_With_Error
End If

If DBEngine.OpenDatabase(dbPath, ReadOnly:=False) Is Nothing Then
    errorString = errorString & vbCr & "Database path " & Chr(34) & dbPath & Chr(34) & " is invalid."
    hasError = True
End If

If hasError = True Then
    GoTo Exit_With_Error
End If


On Error GoTo Exit_With_Error
Set thisDB = Application.CurrentDb
Set inputDB = DBEngine.OpenDatabase(dbPath, ReadOnly:=False)
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
    If Access_Macro_Utilities.CheckObjectExists(importQueries(i, 0), 1) = True Then
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

ImportNALastYrQueries = ""

Exit Function

Exit_With_Error:
  If Not (inputDB Is Nothing) Then
    inputDB.Close
  End If
  Set thisDB = Nothing
  Set inputDB = Nothing
  Set currQuery = Nothing

  If hasError = True Then
    ImportNALastYrQueries = errorString
  Else
    ImportNALastYrQueries = errorString & vbCr & Err.Description
  End If

End Function

