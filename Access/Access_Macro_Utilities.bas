Attribute VB_Name = "Access_Macro_Utilities"
Option Compare Database
Option Explicit
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

' Function goes here if does not exist.
ExitFunc:
    CheckFieldExists = False

End Function
Function CheckObjectExists(objectName As String, objectCode As Integer) As Boolean
'' IN PROGRESS: need to figure out how to check for macros.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Checks if object with name exists in current database.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Note: object type integers are: Table = 0, Query = 1, Form = 2, Report = 3, Macro = 4, Module = 5
If objectCode < 0 Or objectCode > 5 Then
    GoTo Exit_Func_False
End If

On Error GoTo Exit_Func_False
If objectCode = 0 Then
    Dim table As DAO.TableDef
    Set table = CurrentDb.TableDefs(objectName)
    CheckObjectExists = True
ElseIf objectCode = 1 Then
    Dim query As DAO.QueryDef
    Set query = CurrentDb.QueryDefs(objectName)
    CheckObjectExists = True
ElseIf objectCode = 2 Then
    Dim form_Test As form
    Set form_Test = Forms(form_Test)
    CheckObjectExists = True
ElseIf objectCode = 3 Then
    Dim report_Test As report
    Set report_Test = Reports(objectName)
    CheckObjectExists = True
ElseIf objectCode = 4 Then
    '' TODO: determine how to check for macros.
    CheckObjectExists = False
ElseIf objectCode = 5 Then
    Dim module_Test As Module
    Set module_Test = Modules(objectName)
    CheckObjectExists = True
End If

Exit Function

Exit_Func_False:
    CheckObjectExists = False

End Function

Sub OptimizeCodeSettings(turnOff As Boolean)
''''''''' IN PROGRESS: Add more settings.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: turn off settings that make Access VBA less efficient.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If turnOff = True Then
    DoCmd.SetWarnings False
    Application.Echo 0
Else
    DoCmd.SetWarnings True
    Application.Echo 1
End If
End Sub

Function CheckObjectOpen(name As String, objectTypeInt As Integer) As Boolean
'''' IN PROGRESS:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Check if object of type is open.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Note: object type integers are: Table = 0, Query = 1, Form = 2, Report = 3, Macro = 4, Module = 5
'' TODO: add graphical dropdown for corresponding codes and description

If objectTypeInt < 0 Or objectTypeInt > 5 Then
    GoTo Exit_With_Error
End If

On Error GoTo Exit_With_Error
    If SysCmd(acSysCmdGetObjectState, objectTypeInt, name) <> 0 Then
        CheckObjectOpen = True
        Exit Function
    Else
        CheckObjectOpen = False
        Exit Function
    End If
Exit_With_Error:
    CheckObjectOpen = False

End Function

Function CheckDatabaseOpen(ByRef db_in As DAO.Database) As Boolean
'' Description: Returns true if database is open. False otherwise.
On Error GoTo Exit_With_Error


End Function
Sub TurnSettingsBackOn()

Call Access_Macro_Utilities.OptimizeCodeSettings(False)


End Sub
