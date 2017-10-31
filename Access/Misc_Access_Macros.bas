Attribute VB_Name = "Misc_Access_Macros"
Sub RemoveColumnUnderscores()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Removes underscore in column name if not representing a space, otherwise replaces with a space.
'' Commonly happens with the newline character in the column header in Excel.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
        If i = strLen + 1 Then
            ' Do nothing since underscore not present.
        ElseIf StrComp(Mid(attrName, i - 1, 1), " ") = 0 Then
            ' If previous character is a space then is not serving as a placeholder for spaces, thus remove:
            currField.name = Replace(currField.name, "_", "")
        ElseIf IsPunct(Mid(currField.name, 1, 1)) = True Then
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Sets all columns in database that are numeric to Doubles with precision of 2 decimal places.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: replaces all attributes with NULL (string) as value with spaces (true nulls).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Function Run_All_Queries()
'' Description: Run all queries in current database.
Dim timer As Integer: timer = 3
' TODO: potentially loop to run all, add selection criteria (ex: don't run any SELECT queries, only updates).
Dim activeDB As DAO.Database: Set activeDB = CurrentDb
Dim currTable As DAO.TableDef
Dim currQuery As DAO.QueryDef
' Close all tables:
For Each currTable In activeDB.TableDefs
    DoCmd.Close acTable, currTable.name
Next currTable

' Turn off inefficient settings:
Call Access_Macro_Utilities.OptimizeCodeSettings(True)


On Error GoTo Run_All_Queries_Err
' Execute every query in the current database (EXCEPT the XX and TT queries)
For Each currQuery In activeDB.QueryDefs
    If InStr(1, currQuery.name, "XX") = 0 And InStr(1, currQuery.name, "TT") = 0 Then
        ' Display current query name then execute:
        CreateObject("WScript.Shell").PopUp currQuery.name, timer, "Current Query"
        DoCmd.OpenQuery currQuery.name, acViewNormal, acEdit
    End If
Next currQuery

Run_All_Queries_Exit:
    Call Access_Macro_Utilities.OptimizeCodeSettings(False)
    Exit Function

Run_All_Queries_Err:
    MsgBox "Query failed at " & currQuery.name
    Call Access_Macro_Utilities.OptimizeCodeSettings(False)
    Resume Run_All_Queries_Exit

End Function

Sub ExportSQLQueriesToTextFile()
'' Description: Write all queries in current database in SQL form to file. Outputs to specified location.
'' TODO: use regular expressions to determine if element of SELECT clause (ex [TableName].[Column Name])

' Turn off inefficient settings:
Call Access_Macro_Utilities.OptimizeCodeSettings(True)

Dim activeDB As DAO.Database: Set activeDB = CurrentDb
Dim currQuery As DAO.QueryDef
Dim filepathString As String: filepathString = InputBox("Enter filepath for queries: ")
Dim startIndex As Integer: startIndex = InStrRev(activeDB.name, "\") + 1
filepathString = filepathString & "\" & Mid(activeDB.name, startIndex, InStr(1, activeDB.name, ".") - startIndex) & "_Queries.txt"
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
'' Todo: add pattern matching with regExp (https://stackoverflow.com/questions/37443992/implementing-regex-into-access-vba-for-password-complexity)

On Error GoTo File_Error
' Generate the blank text file to write to:
Dim Fileout As Object: Set Fileout = fso.CreateTextFile(filepathString, True, True)



' Write the queries in SQL form to file, with corresponding title in Access database:
Dim formattedQueryString, remainingQuery, tempString As String
Dim keywordIndex, index As Integer
Dim sqlKeywords() As String: sqlKeywords = Split("FROM,WHERE,GROUP,HAVING,INNER,OUTER,JOIN,ON,SET", ",")
Dim hasChanged As Boolean

On Error GoTo Loop_Error

For Each currQuery In activeDB.QueryDefs
    ' Write the query name used in Access:
    Fileout.Write "----------" & currQuery.name & "----------" & vbNewLine
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
        Fileout.Write formattedQueryString & vbNewLine
    Else
        Fileout.Write currQuery.SQL & vbNewLine
    End If
Next currQuery

' Close file and zero out references:
Fileout.Close
Set Fileout = Nothing
Set activeDB = Nothing
Set currQuery = Nothing
Set fso = Nothing
Exit Sub

' Subroutine goes here on file error.
File_Error:
    MsgBox "Error: file path is invalid."
    Call Access_Macro_Utilities.OptimizeCodeSettings(False)
    Set Fileout = Nothing
    Set activeDB = Nothing
    Set currQuery = Nothing
    Set fso = Nothing
    Exit Sub
' Subroutine goes here on loop error.
Loop_Error:
    MsgBox "Error: loop error."
    Call Access_Macro_Utilities.OptimizeCodeSettings(False)
    Set Fileout = Nothing
    Set activeDB = Nothing
    Set currQuery = Nothing
    Set fso = Nothing
    Exit Sub
    
End Sub
Sub ImportSQLQueriesFromFile()
''''''''''''''''''' IN PROGRESS
'' Description: Import SQL queries from text file, convert into queries with Access.




End Sub


