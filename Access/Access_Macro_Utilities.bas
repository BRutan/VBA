Attribute VB_Name = "Module1"
Sub RemoveColumnUnderscores()
'' Description: Removes underscore in column name if not representing a space, otherwise replaces with a space.
'' Commonly happens with the newline character in the column header in Excel.

Dim activeDB As DAO.Database: Set activeDB = CurrentDb
Dim currTableDef As DAO.TableDef
Dim currField As DAO.field

Dim i, strLen As Integer
Dim attrName As String
Dim asSpace As Boolean

For Each currTableDef In activeDB.TableDefs
    attrName = currTableDef.name
' Determine if underscore is serving as replacement for space character, if true then replace with space. Otherwise remove.
    For Each currField In currTableDef.Fields
        i = 1
        attrName = currField.name
        strLen = Len(attrName)
        asSpace = False
        ' Find index of underscore:
        Do While StrComp(Mid(attrName, i, 1), "_") <> 0 And i <> strLen + 1
            i = i + 1
        Loop
        ' Make appropriate changes to attribute name if necessary:
        If i = strLen Then
            ' Do nothing since underscore not present.
        ElseIf StrComp(Mid(currField.name, i - 1, 1), " ") = 0 Then
            ' If previous character is a space then is not serving as a placeholder for spaces, thus remove:
            currField.name = Replace(currField.name, "_", "")
        ElseIf i <> 1 And IsPunct(Mid(currField.name, i - 1, 1)) = True Then
        ' If previous character is punctuation assume that is not a placeholder (like CPT/_HCPCS):
            currField.name = Replace(currField.name, "_", "")
        Else
        ' Otherwise assume serving as a space placeholder:
            currField.name = Replace(currField.name, "_", " ")
        End If
    Next currField
Next currTableDef

Set activeDB = Nothing
Set currTableDef = Nothing
Set currField = Nothing

End Sub
Sub SetNumberFieldsToDoubles()
'' Description: Sets all columns in database that are numeric to Doubles with precision of 2 decimal places.

Dim activeDB As DAO.Database: Set activeDB = CurrentDb
Dim currTableDef As DAO.TableDef
Dim currField As DAO.field

' If type is 3 - 7, then assume is number type, thus set to Double with 2
' decimals precision.

For Each currTableDef In activeDB.TableDefs
    For Each currField In currTableDef.Fields
        If currField.Type >= 3 And currField.Type <= 7 Then
            activeDB.Execute "ALTER TABLE [" & currTableDef.name & "] ALTER COLUMN [" & currField.name & "] DOUBLE"
        End If
    Next currField
Next currTableDef

Set activeDB = Nothing
Set currTableDef = Nothing
Set currField = Nothing

End Sub
Sub ClearNulls()
'' Description: replaces all attributes with NULL (string) as value to Null (true nulls).

Dim activeDB As DAO.Database: Set activeDB = CurrentDb
Dim currTable As DAO.TableDef
Dim currField As DAO.field

' Iterate through each column in each table in the current database:
For Each currTable In activeDB.TableDefs
    If InStr(1, currTable.name, "MSys") = 0 Then
        For Each currField In currTable.Fields
            ' Execute update queries, setting every attribute where 'NULL' to Null (true nulls).
            If currField.Type = dbText Then
                activeDB.Execute "UPDATE [" & currTable.name & "] SET [" & currField.name & "] = Null WHERE [" & currField.name & "] = 'NULL'"
            End If
        Next currField
    End If
Next currTable

MsgBox "Done replacing 'NULL's with Null."

End Sub
Sub AppendNAPriceDiffFields()
'' Description:
' Note: to work, must set the Net Analysis table name to "Net Analysis".

Dim activeDB As DAO.Database: Set activeDB = CurrentDb
' Skip to the ExitSub label if the 'Net Analysis' table does not exist, or some other error occurs:
On Error GoTo ExitSub
Dim naTable As DAO.TableDef: Set naTable = activeDB.TableDefs("Net Analysis")
Dim currTableDef2 As DAO.TableDef
Dim currField As DAO.field
Dim nameString As String
Dim needToAddLast3 As Boolean: needToAddLast3 = True

Dim namesArray() As String
Dim currArrayIndex As Integer: currArrayIndex = 0

' Generate the table update fields based upon tables from Rate Matrix:
For Each currTableDef2 In activeDB.TableDefs
    ' Get all characters to the right up dashes:
    If InStr(currTableDef2.name, "IP") <> 0 Or InStr(currTableDef2.name, "OP") <> 0 Then
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
    If CheckFieldExists(naTable.name, namesArray(currArrayIndex)) = True Then
        naTable.Fields.Delete namesArray(currArrayIndex)
    End If
Next currArrayIndex

' Append columns:
For currArrayIndex = 0 To UBound(namesArray)
    naTable.Fields.Append naTable.CreateField(namesArray(currArrayIndex), dbDouble)
Next currArrayIndex

' Nothing out all the referenced objects and exit:
Set activeDB = Nothing
Set naTable = Nothing
Set currTableDef2 = Nothing
Exit Sub
' Subroutine goes here if error occurs.
ExitSub:
    MsgBox "Error: 'Net Analysis' table does not exist OR Net Analysis table is open."
    Exit Sub

End Sub
Function Run_All_Queries()
'' Description: Run all queries in current database.
'' Todo: Add output failed query to file.
Dim startQuery As String: startQuery = Trim(InputBox("Enter the query number to start Run-All Macro:"))
Dim queriesToRun() As String
Dim queryExists As Boolean
Dim timer As Integer: timer = 3
Dim activeDB As DAO.Database: Set activeDB = CurrentDb
Dim currTable As DAO.TableDef
Dim currQuery As DAO.QueryDef

'' Regular Expression:
Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
regExp.Pattern = "[0-9]?[0-9][a-z]?%"
regExp.Global = True
' Close all tables:
For Each currTable In activeDB.TableDefs
    DoCmd.Close acTable, currTable.name
Next currTable

On Error GoTo Macro_Err
' Turn off warnings:
DoCmd.SetWarnings False
' Determine if starting query was entered, or invalid.
If regExp.Test(startQuery) = False Then
    MsgBox "Query is invalid/not entered. Running all queries."
    GoTo Run_All
Else
    ' Search for query among query names in db:
    Dim queriesToRunIndex As Integer: queriesToRunIndex = 1
    For Each currQuery In activeDB.QueryDefs
        If InStr(1, currQuery.name, startQuery & " -") <> 0 Or InStr(1, currQuery.name, startQuery & "-") <> 0 Then
            queryExists = True
            ReDim queriesToRun(0): queriesToRun(0) = currQuery.name
        ElseIf queryExists = True And InStr(1, currQuery.name, "XX") = 0 And InStr(1, currQuery.name, "TT") = 0 Then
            ReDim Preserve queriesToRun(queriesToRunIndex)
            queriesToRun(queriesToRunIndex) = currQuery.name
            queriesToRunIndex = queriesToRunIndex + 1
        End If
    Next currQuery
    If queryExists = False Then
        MsgBox "Query does not exist. Running all queries."
        GoTo Run_All
    End If
    GoTo Run_Starting_At
End If

'''''''''''''' Option 1:
' Run all queries option:
Run_All:
' Execute every query in the current database (EXCEPT the XX and TT queries)
For Each currQuery In activeDB.QueryDefs
    If InStr(1, currQuery.name, "XX") = 0 And InStr(1, currQuery.name, "TT") = 0 Then
        ' Display current query name then execute:
        CreateObject("WScript.Shell").PopUp currQuery.name, timer, "Current Query"
        DoCmd.OpenQuery currQuery.name, acViewNormal, acEdit
    End If
Next currQuery
GoTo Macro_Exit

Run_Starting_At:
'''''''''''''' Option 2:
' Run query starting at provided query option:
Dim currIter As Variant
For Each currIter In queriesToRun
    CreateObject("WScript.Shell").PopUp currIter, timer, "Current Query"
    DoCmd.OpenQuery currIter, acViewNormal, acEdit
Next currIter
GoTo Macro_Exit

' GOTO Labels:
Macro_Exit:
    DoCmd.SetWarnings True
    Exit Function

Macro_Err:
    ' Export last executed query to text file in local folder:
    Dim fileOut As Object: Set fileOut = CreateObject("Scripting.FileSystemObject")
    Dim file As Object: Set file = fileOut.CreateTestFile(Application.CurrentProject.Path & "\" & CurrentDb.name & "_Messages", True)
    ' If currQuery is empty then the Run_Starting_At option must have been used:
    If currQuery Is Nothing Then
        file.WriteLine Now & ": Last query was " & currIter
    Else
        file.WriteLine Now & ": Last query was " & currQuery.name
    End If
    DoCmd.SetWarnings True
    Resume Macro_Exit

End Function
Private Function IsPunct(ByVal char As String) As Boolean
'' Description: Determines if passed character is a punctuation character.
If Asc(Mid(char, 1, 1)) = 33 Or Asc(Mid(char, 1, 1)) = 40 Or Asc(Mid(char, 1, 1)) = 47 Or Asc(Mid(char, 1, 1)) = 59 _
    Or Asc(Mid(char, 1, 1)) = 63 Or Asc(Mid(char, 1, 1)) = 92 Then
        IsPunct = True
    Else
        IsPunct = False
End If

End Function
Function CheckFieldExists(tableDefName As String, name As String) As Boolean
'' Description: Checks if field with passed name is already present in the passed table (in current database).
On Error GoTo ExitFunc
Dim field As DAO.field

    Set field = CurrentDb.TableDefs(tableDefName).Fields(name)
    ' Skipped if the field name does not exist in the table.
    CheckFieldExists = True
    Exit Function
' Function goes
ExitFunc:
    CheckFieldExists = False

End Function

Sub ExportSQLQueriesToTextFile()
'' Description: Write all queries in current database in SQL form to file. Outputs to specified location.
'' TODO: use regular expressions to determine if element of SELECT clause (ex [TableName].[Column Name])
Dim activeDB As DAO.Database: Set activeDB = CurrentDb
Dim currQuery As DAO.QueryDef
Dim filepathString As String: filepathString = InputBox("Enter filepath for queries: ")
Dim startIndex As Integer: startIndex = InStrRev(activeDB.name, "\") + 1
filepathString = filepathString & "\" & Mid(activeDB.name, startIndex, InStr(1, activeDB.name, ".") - startIndex) & "_Queries.txt"
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
'' Todo: add pattern matching with regExp (https://stackoverflow.com/questions/37443992/implementing-regex-into-access-vba-for-password-complexity)
''   to add <tab> plus line break on commas.
On Error GoTo File_Error
' Generate the blank text file to write to:
Dim fileOut As Object: Set fileOut = fso.CreateTextFile(filepathString, True, True)

' Write the queries in SQL form to file, with corresponding title in Access database:
Dim formattedQueryString, remainingQuery, tempString As String
Dim keywordIndex, index As Integer
Dim sqlKeywords() As String: sqlKeywords = Split("FROM,WHERE,GROUP,HAVING,INNER,OUTER,JOIN,ON,SET", ",")
Dim hasChanged As Boolean

On Error GoTo Loop_Error

For Each currQuery In activeDB.QueryDefs
    ' Write the query name used in Access:
    fileOut.Write "----------" & currQuery.name & "----------" & vbNewLine
    remainingQuery = currQuery.SQL
    formattedQueryString = vbNullString
    tempString = vbNullString
    keywordIndex = 0
    ' Format the query string:
    For index = 0 To UBound(sqlKeywords)
        If InStr(Len(tempString) + 1, remainingQuery, sqlKeywords(index)) <> 0 Then
            keywordIndex = InStr(1, remainingQuery, sqlKeywords(index))
            tempString = Mid(remainingQuery, 1, keywordIndex - 1) & vbNewLine
            formattedQueryString = formattedQueryString & tempString
            remainingQuery = Mid(remainingQuery, keywordIndex, Len(remainingQuery) - Len(tempString) + 1)
            hasChanged = True
        End If
    Next index
    ' Append remaining query to formatted query string, then write formatted query to file:
    If hasChanged = True Then
        formattedQueryString = formattedQueryString & remainingQuery
        fileOut.Write formattedQueryString & vbNewLine
    Else
        fileOut.Write currQuery.SQL & vbNewLine
    End If
Next currQuery

' Close file and zero out references:
fileOut.Close
Set fileOut = Nothing
Set activeDB = Nothing
Set currQuery = Nothing
Set fso = Nothing
Exit Sub

' Subroutine goes here on file error.
File_Error:
    MsgBox "Error: file path is invalid."
    Set fileOut = Nothing
    Set activeDB = Nothing
    Set currQuery = Nothing
    Set fso = Nothing
    Exit Sub
' Subroutine goes here on loop error.
Loop_Error:
    MsgBox "Error: loop error."
    Set fileOut = Nothing
    Set activeDB = Nothing
    Set currQuery = Nothing
    Set fso = Nothing
    Exit Sub
    
End Sub
Sub SpawnUpdateQueries()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' IN PROGRESS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: generate the update queries using the numbered tables in the current database.
'' Note: Requires that Rate Matrix tables be loaded.


'''' TODO: use the SearchForColumn function to return appropriate columns, instead of doing manually.

Dim RE1, RE2 As New regExp
RE1.Pattern = "%[0-9]?[0-9][a-z]?%"
RE1.Global = True
RE2.Pattern = "Hospital[/Model]"
RE2.Global = True

Dim activeDB As DAO.Database: Set activeDB = CurrentDb
Dim currTable, naTable As DAO.TableDef
Dim currColumn, naHospitalColumn As DAO.field
Dim hasHospitalColumn As Boolean
Dim queryName, SQLQuery As String
Dim errorMessage As String

' Find the NA Table:
For Each currTable In activeDB.TableDefs
    If InStr(1, currTable.name, "Net Analysis") <> 0 Then
        Set naTable = currTable
        Exit For
    End If
Next currTable

' Determine if AppendNAPriceDiffFields macro needs to be run via searching the Net Analysis table:
Call Module1.AppendNAPriceDiffFields ' If already present, skips.

' Find the hospital/model column in the Net Analysis table
For Each currColumn In naTable.Fields
    If RE2.Test(currColumn.name) = True Then
        Set naHospitalColumn = currColumn
        hasHospitalColumn = True
        Exit For
    End If
Next currColumn

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Main Routine
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim codeTypes As New StringVector
Dim pmtMethod As String
Dim regExpColSearch As regExp
Dim validityCheck As Boolean

Dim searchColumns(10) As String

'' TODO: refactor so append to query using single (or multiple) arrays grouped by purpose.
For Each currTable In activeDB.TableDefs
    codeTypes.Clear
    validityCheck = False
    ' Generate query if table name matching regular expression exists (01-XX with optional lowercase letter).
    '''' TODO: run regexp test to find the Plan Name column.
    If RE.Test(currTable.name) = True Then
        queryName = currTable.name & " Updates"
        SQLQuery = "UPDATE [" & naTable.name & " AS A INNER JOIN [" & currTable.name & "] AS B ON A.[FY " & dbYear & " Plan Name] = B.[FY " & dbYear & " Plan Name] AND "
        ' Search for Hospital/Model column in the current query table:
        If hasHospitalColumn = True Then
            For Each currColumn In currTable.Fields
                If RE2.Test(currColumn.name) = True Then
                    hasHospitalColumn = True
                    SQLQuery = SQLQuery & "A.[" & naHospitalColumn.name & "] = B.[" & currColumn.name & "]"
                    Exit For
                End If
            Next currColumn
        End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Search for identifier types:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' TODO: refactor to run in one loop.
        If InStr(1, currTable.name, "TBG") = 0 Then
        SQLQuery = SQLQuery & " AND "
        ''''''''' Determine if link on Rev Code:
            If InStr(1, currTable.name, "RC") <> 0 Then
                ' Search for column similar to "Rev Code" in Net Analysis table and current table:
                For Each currColumn In naTable.Fields
                    If InStr(1, currColumn.name, "Rev ") <> 0 Then
                        ' Append link on net analysis column:
                        validityCheck = True
                        SQLQuery = SQLQuery & "A.[" & currColumn.name & "] = "
                        Exit For
                    End If
                Next currColumn
                ' If no linkable Rev Code column present, display error message and exit macro:
                If validityCheck = False Then
                    errorMessage = "Error: No Rev Code column present in " & naTable.name & ". Exiting macro."
                    GoTo Exit_With_Error
                End If
                validityCheck = False
                For Each currColumn In currTable.Fields
                    If InStr(1, currColumn.name, "Rev Code") <> 0 Then
                        ' Append link on current table:
                        validityCheck = True
                        SQLQuery = SQLQuery & "B.[" & currColumn.name & "]"
                        Exit For
                    End If
                Next currColumn
                ' If no Rev Code column in current table, display error message and exit macro:
                If validityCheck = False Then
                    errorMessage = "Error: No Rev Code column present in " & currTable.name & ". Exiting macro."
                    GoTo Exit_With_Error
                End If
            End If
        ''''''''' Determine if link on CPT/HCPCS:
            If InStr(1, currTable.name, "CPT") <> 0 Then
                SQLQuery = SQLQuery & " AND "
                '' Search for CPT column in the Net Analysis Table:
                For Each currColumn In naTable.Fields
                    If InStr(1, currColumn.name, "CPT") <> 0 Then
                        validityCheck = True
                        SQLQuery = SQLQuery & "A.[" & currColumn.name & "] = "
                        Exit For
                    End If
                Next currColumn
                ' If no CPT column present, display error message and exit macro:
                If validityCheck = False Then
                    errorMessage = "Error: No CPT column present in " & currTable.name & ". Exiting macro."
                    GoTo Exit_With_Error
                End If
                validityCheck = False
                '' Search for CPT column in current table:
                For Each currColumn In currTable.Fields
                    If InStr(1, currColumn.name, "CPT") <> 0 Then
                        validityCheck = True
                        SQLQuery = SQLQuery & "B.[" & currColumn.name & "]"
                        Exit For
                    End If
                Next currColumn
                ' If no CPT column in current table, display error message and exit macro:
                If validityCheck = False Then
                    errorMessage = "Error: No CPT column present in " & currTable.name & ". Exiting macro."
                    GoTo Exit_With_Error
                End If
            End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Get payment type (PIA, PP, % of Charge) and append query accordingly:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If InStr(1, currTable.name, "%") <> 0 Then
                ' Determine if IP or OP:
                If InStr(1, currTable.name, "IP") <> 0 Then
                    ' Is inpatient, thus set on the [IP %] column:
                    SQLQuery = SQLQuery & " SET "
                ElseIf InStr(1, currTable.name, "OP") <> 0 Then
                    ' Is outpatient, thus set on the [OP %] column:
                Else
                    ' Table is invalid, throw error and exit macro:
                    errorMessage = "Error: table " & currTable.name & " is invalid."
                    GoTo Exit_With_Error
                End If
            ElseIf InStr(1, currTable.name, "PP") <> 0 Then
                ' Determine if IP or OP:
                If InStr(1, currTable.name, "IP") <> 0 Then
                    ' Is inpatient, thus set on the [IP %] column:
                ElseIf InStr(1, currTable.name, "OP") <> 0 Then
                    ' Is outpatient, thus set on the [OP %] column:
                Else
                    ' Table is invalid, throw error and exit macro:
                    errorMessage = "Error: table " & currTable.name & " is invalid."
                    GoTo Exit_With_Error
                End If
            ElseIf InStr(1, currTable.name, "PIA") <> 0 Then
                ' Determine if IP or OP:
                If InStr(1, currTable.name, "IP") <> 0 Then
                    ' Is inpatient, thus set on the [IP %] column:
                ElseIf InStr(1, currTable.name, "OP") <> 0 Then
                    ' Is outpatient, thus set on the [OP %] column:
                Else
                    ' Table is invalid, throw error and exit macro:
                    errorMessage = "Error: table " & currTable.name & " is invalid."
                    GoTo Exit_With_Error
                End If
            Else
            ' Table is invalid
            End If
            
        ElseIf InStr(1, currTable.name, "TBG") <> 0 Then
        ' Determine if IP or OP:
            If InStr(1, currTable.name, "IP") <> 0 Then
                ' Is inpatient, thus set on the [IP %] column:
            ElseIf InStr(1, currTable.name, "OP") <> 0 Then
                ' Is outpatient, thus set on the [OP %] column:
            Else
                ' Table is invalid, throw error and exit macro:
                errorMessage = "Error: table " & currTable.name & " is invalid."
                GoTo Exit_With_Error
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Generate the query using the completed SQLQuery string and table name:
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    End If
Next currTable

Exit Sub

Exit_With_Error:
    MsgBox errorMessage
    DoCmd.SetWarnings = True
    Exit Sub


End Sub

Sub ResetWarningsOn()
'' Description: turn warnings back on (overcome issues when warnings are not reset following query abrupt stop).

DoCmd.SetWarnings True

End Sub
