Attribute VB_Name = "HotkeyedMacros"
Option Explicit

Sub ConvertRangeToNumber()
Attribute ConvertRangeToNumber.VB_Description = "Converts string values in selected range to numbers. Useful for ensuring VLOOKUP works properly."
Attribute ConvertRangeToNumber.VB_ProcData.VB_Invoke_Func = "N\n14"

' Currently assigned to CTRL + SHIFT + N.
Call Macro_Utilities.CodeOptimizeSettings(True)

Selection.NumberFormat = "General"
Selection.value = Selection.value

Call Macro_Utilities.CodeOptimizeSettings(False)

End Sub

Sub OptimalMonetaryFormat()

' Currently assigned to CTRL + SHIFT + M.
Call Macro_Utilities.CodeOptimizeSettings(True)

On Error GoTo Leave

Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* " - "??_);_(@_);"
Selection.value = Selection.value

Leave:
    Call Macro_Utilities.CodeOptimizeSettings(True)

End Sub

