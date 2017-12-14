Attribute VB_Name = "Benchmarking"
Option Explicit
Function CaseStatementCreator(onColumn As String, InputRange As Range, valueRange As Range, withDelimits As Boolean) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Converts a series of text if statements on passed range into a case statement to be used in a SQL SELECT statement.

' Input checking:
If InputRange.Rows.count <> valueRange.Rows.count Or InputRange.Columns.count > 1 Or valueRange.Columns.count > 1 Then
    CaseStatementCreator = "#N/A"
    Exit Function
End If

Dim outString As String: outString = "CASE " & onColumn
Dim row As Integer: row = 0
Dim delimit As String
If withDelimits = True Then
    delimit = "'"
Else
    delimit = vbNullString
End If

For row = 1 To InputRange.Rows.count
    outString = outString & " WHEN " & delimit & InputRange.cells(row, 1).value & delimit & _
        " THEN " & delimit & valueRange.cells(row, 1).value & delimit
Next row

outString = outString & " END "

CaseStatementCreator = outString

End Function
Function Distance_Latitude_Longitude(latLong1 As String, latLong2 As String) As Double
'' Description: Find distance between two latitude/longitude data in the form of "DD°MM'DIR / DD°MM'DIR"
'''''' TODO:
''' 1. Double check distance formula
''' 2. Add ability to determine if seconds and greater precision added to string.

'' Convert minutes/seconds notation to decimal degree:
latLong1 = Trim(latLong1)
latLong2 = Trim(latLong2)

On Error GoTo Exit_With_Error

Dim lat1 As Double: lat1 = CDbl(Mid(latLong1, 1, InStr(1, latLong1, "°") - 1))
latLong1 = Mid(latLong1, InStr(1, latLong1, "°") + 1, Len(latLong1))
lat1 = lat1 + CDbl(Mid(latLong1, 1, InStr(1, latLong1, "'") - 1)) / 60
latLong1 = Trim(Mid(latLong1, InStr(1, latLong1, "/") + 1, Len(latLong1)))

Dim lon1 As Double: lon1 = CDbl(Mid(latLong1, 1, InStr(1, latLong1, "°") - 1))
latLong1 = Mid(latLong1, InStr(1, latLong1, "°") + 1, Len(latLong1))
lon1 = lon1 + CDbl(Mid(latLong1, 1, InStr(1, latLong1, "'") - 1)) / 60

Dim lat2 As Double: lat2 = CDbl(Mid(latLong2, 1, InStr(1, latLong2, "°") - 1))
latLong2 = Mid(latLong2, InStr(1, latLong2, "°") + 1, Len(latLong2))
lat2 = lat2 + CDbl(Mid(latLong2, 1, InStr(1, latLong2, "'") - 1)) / 60
latLong2 = Trim(Mid(latLong2, InStr(1, latLong2, "/") + 1, Len(latLong2)))

Dim lon2 As Double: lon2 = CDbl(Mid(latLong2, 1, InStr(1, latLong2, "°") - 1))
latLong2 = Mid(latLong2, InStr(1, latLong2, "°") + 1, Len(latLong2))
lon2 = lon2 + CDbl(Mid(latLong2, 1, InStr(1, latLong2, "'") - 1)) / 60

'' Convert lat and lon degrees to radians:
Dim dLat As Double: dLat = (lat2 - lat1) / 180 * Application.WorksheetFunction.Pi
Dim dLon As Double: dLon = (lon2 - lon1) / 180 * Application.WorksheetFunction.Pi

'' Calculate distance using haversine method:
Dim a As Double: a = Sin(dLat / 2) * Sin(dLat / 2) + Cos(lat1 / 180 * Application.WorksheetFunction.Pi) _
                    * Cos(lat2 / 180 * Application.WorksheetFunction.Pi) * Sin(dLon / 2) * Sin(dLon / 2)
Dim c As Double: c = 2 * Application.WorksheetFunction.Atan2(Sqr(a), Sqr((1 - a)))
Dim d As Double: d = 6371 / 1.609 * c

'' Return distance:

Distance_Latitude_Longitude = d

Exit Function

Exit_With_Error:
    Distance_Latitude_Longitude = -99
    Exit Function

End Function

Public Function GetDistanceCoord(latLong1 As String, latLong2 As String, unit As String) As Double
'' Description:

    Dim lat1 As Double: lat1 = CDbl(Mid(latLong1, 1, InStr(1, latLong1, "°") - 1))
    latLong1 = Mid(latLong1, InStr(1, latLong1, "°") + 1, Len(latLong1))
    lat1 = lat1 + CDbl(Mid(latLong1, 1, InStr(1, latLong1, "'") - 1)) / 60
    latLong1 = Trim(Mid(latLong1, InStr(1, latLong1, "/") + 1, Len(latLong1)))
    
    Dim lon1 As Double: lon1 = CDbl(Mid(latLong1, 1, InStr(1, latLong1, "°") - 1))
    latLong1 = Mid(latLong1, InStr(1, latLong1, "°") + 1, Len(latLong1))
    lon1 = lon1 + CDbl(Mid(latLong1, 1, InStr(1, latLong1, "'") - 1)) / 60
    
    Dim lat2 As Double: lat2 = CDbl(Mid(latLong2, 1, InStr(1, latLong2, "°") - 1))
    latLong2 = Mid(latLong2, InStr(1, latLong2, "°") + 1, Len(latLong2))
    lat2 = lat2 + CDbl(Mid(latLong2, 1, InStr(1, latLong2, "'") - 1)) / 60
    latLong2 = Trim(Mid(latLong2, InStr(1, latLong2, "/") + 1, Len(latLong2)))
    
    Dim lon2 As Double: lon2 = CDbl(Mid(latLong2, 1, InStr(1, latLong2, "°") - 1))
    latLong2 = Mid(latLong2, InStr(1, latLong2, "°") + 1, Len(latLong2))
    lon2 = lon2 + CDbl(Mid(latLong2, 1, InStr(1, latLong2, "'") - 1)) / 60

    Dim theta As Double: theta = lon1 - lon2
    Dim dist As Double: dist = Math.Sin(deg2rad(lat1)) * Math.Sin(deg2rad(lat2)) + Math.Cos(deg2rad(lat1)) * Math.Cos(deg2rad(lat2)) * Math.Cos(deg2rad(theta))
    dist = WorksheetFunction.Acos(dist)
    dist = rad2deg(dist)
    dist = dist * 60 * 1.1515
    If unit = "K" Then
        dist = dist * 1.609344
    ElseIf unit = "N" Then
        dist = dist * 0.8684
    End If
    GetDistanceCoord = dist
End Function
Public Function GetDistance(start As String, dest As String, units As String)

    Call Macro_Utilities.CodeOptimizeSettings(True)
    
' Description:
    Dim firstVal As String, secondVal As String, lastVal As String
    Dim objHTTP As Object
    Dim Url As String
    Dim regex As regExp
    Dim matches As Object
    Dim tmpVal As String
    On Error GoTo Exit_With_Error
    firstVal = "http://maps.googleapis.com/maps/api/distancematrix/json?origins="
    secondVal = "&destinations="
    lastVal = "&mode=car&language=pl&sensor=false"
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Url = firstVal & Replace(start, " ", "+") & secondVal & Replace(dest, " ", "+") & lastVal
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    If InStr(objHTTP.responseText, """distance"" : {") = 0 Then GoTo Exit_With_Error
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = """value"".*?([0-9]+)": regex.Global = False
    Set matches = regex.Execute(objHTTP.responseText)
    tmpVal = Replace(matches(0).SubMatches(0), ".", Application.International(xlListSeparator))
    If StrComp(units, "M") = 0 Or StrComp(units, "m") = 0 Then
        ' Return number of miles between addresses (query returns meters):
        GetDistance = CDbl(tmpVal) * 0.000621371
    ElseIf StrComp(units, "K") = 0 Or StrComp(units, "k") = 0 Then
        ' Return number of kilometers between addresses (query returns meters):
        GetDistance = CDbl(tmpVal) / 1000
    Else
        GoTo Exit_With_Error
    End If
    Call Macro_Utilities.CodeOptimizeSettings(False)
    Exit Function
Exit_With_Error:
    Call Macro_Utilities.CodeOptimizeSettings(False)
    GetDistance = -99
End Function

Function CosineSimilarity(ByRef arr1 As ArrayList, ByRef arr2 As ArrayList) As LongLong
'''''''''''''''''''''''''''''''''''
'' IN PROGRESS:
'''''''''''''''''''''''''''''''''''
'' Description: Compute cosine "similarity" of two ArrayLists:

' Append zeros until two arrays have same length:
If arr1.count < arr2.count Then
    Do While arr1.count < arr2.count
        arr1.Add (0)
    Loop
ElseIf arr1.count > arr2.count Then
    Do While arr1.count > arr2.count
        arr2.Add (0)
    Loop
End If

Dim iter As Long
Dim val1, val2 As LongLong

' Perform the cosine similarity calculation routine:
For iter = 0 To arr1.count

Next iter

End Function

Function TrendFactorCalc(selectedRange As Range) As Double
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Calculates trend factor as evenly weighted average geometric mean for each revenue code in passed range.
'' Assumes that each row contains prices.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo Exit_With_Error
Application.Volatile

Dim currRange As Range: Set currRange = selectedRange
Dim row, col As Long
Dim currSum As Double: currSum = 0
Dim currAccum As Double
Dim nonEmptyRows As Long: nonEmptyRows = 0

For row = 1 To currRange.Rows.count
    If StrComp(Trim(currRange.cells(row, 1).value), vbNullString) <> 0 Then
        nonEmptyRows = nonEmptyRows + 1
        currAccum = 1
        ' Accumulate all growth factors for current revenue code:
        For col = 1 To currRange.Columns.count - 1
            currAccum = currAccum * (1 + (CDbl(currRange.cells(row, col + 1).value) / CDbl(currRange.cells(row, col).value) - 1))
        Next col
        ' Add the geometric average change to the accumulation:
        currSum = currSum + (currAccum ^ (1 / (currRange.Columns.count - 1)) - 1)
    End If
Next row

TrendFactorCalc = currSum / CDbl(nonEmptyRows)

Exit Function

Exit_With_Error:
    TrendFactorCalc = -99
    Exit Function
    
End Function




Function deg2rad(ByVal deg As Double) As Double
    deg2rad = (deg * WorksheetFunction.Pi / 180#)
End Function
 
Function rad2deg(ByVal rad As Double) As Double
    rad2deg = rad / WorksheetFunction.Pi * 180#
End Function
