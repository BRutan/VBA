Attribute VB_Name = "Options"
Option Explicit

Function BSMPrice(ByRef T_in As String, ByRef S_0_in As String, ByRef sigma_in As String, ByRef r_in As String, ByRef K_in As String, IsCall As Boolean) As Double
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Calculates Black-Scholes-Merton European option price.

On Error GoTo Exit_With_Error
Application.Volatile
Dim T, S_0, sigma, r, K As Double: T = CDbl(T_in): S_0 = CDbl(S_0_in): sigma = CDbl(sigma_in): r = CDbl(r_in): K = CDbl(K_in)

If T < 0 Or S_0 < 0 Or sigma < 0 Or K < 0 Then
   GoTo Exit_With_Error
End If

Dim d1 As Double: d1 = (Math.Log(S_0 / K) + (r + (sigma * sigma) / 2) * T) / (sigma * Sqr(T))
Dim d2 As Double: d2 = d1 - sigma * Sqr(T)
Dim outputPrice As Double

outputPrice = S_0 * Application.WorksheetFunction.Norm_Dist(d1, 0, 1, True) - K * Math.Exp(-r * T) * Application.WorksheetFunction.Norm_Dist(d2, 0, 1, True)
    
If IsCall = False Then
    outputPrice = outputPrice + K * Math.Exp(-r * T) - S_0
End If

BSMPrice = outputPrice

Exit Function

Exit_With_Error:
    BSMPrice = -99
    Exit Function

End Function

Function Gamma_Option(ByRef T_in As String, ByRef S_0_in As String, ByRef sigma_in As String, ByRef r_in As String, ByRef K_in As String) As Double
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
On Error GoTo Exit_With_Error
Application.Volatile
Dim T, S_0, sigma, r, K As Double: T = CDbl(T_in): S_0 = CDbl(S_0_in): sigma = CDbl(sigma_in): r = CDbl(r_in): K = CDbl(K_in)

If T < 0 Or S_0 < 0 Or sigma < 0 Or K < 0 Then
   GoTo Exit_With_Error
End If

Dim d1 As Double: d1 = (Math.Log(S_0 / K) + (r + (sigma * sigma) / 2) * T) / (sigma * Sqr(T))

Gamma_Option = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, False) / (S_0 * sigma * Sqr(T))
Exit Function

Exit_With_Error:
    Gamma_Option = -99

End Function

Function Vega_Option(ByRef T_in As String, ByRef S_0_in As String, ByRef sigma_in As String, ByRef r_in As String, ByRef K_in As String) As Double
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo Exit_With_Error
Application.Volatile
Dim T, S_0, sigma, r, K As Double: T = CDbl(T_in): S_0 = CDbl(S_0_in): sigma = CDbl(sigma_in): r = CDbl(r_in): K = CDbl(K_in)

If T < 0 Or S_0 < 0 Or sigma < 0 Or K < 0 Then
   GoTo Exit_With_Error
End If

Dim d1 As Double: d1 = (Math.Log(S_0 / K) + (r + (sigma * sigma) / 2) * T) / (sigma * Sqr(T))

Vega_Option = S_0 * Application.WorksheetFunction.Norm_Dist(d1, 0, 1, False) * Sqr(T)

Exit Function

Exit_With_Error:
    Vega_Option = -99

End Function

Function Rho_Option(ByRef T_in As Range, ByRef S_0_in As Range, ByRef sigma_in As Range, ByRef r_in As Range, ByRef K_in As Range, IsCall As Boolean) As Double
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Calculate Rho (sensitivity of option price to change in risk free rate) using passed parameters.

On Error GoTo Exit_With_Error
Application.Volatile
Dim T, S_0, sigma, r, K As Double: T = CDbl(T_in): S_0 = CDbl(S_0_in): sigma = CDbl(sigma_in): r = CDbl(r_in): K = CDbl(K_in)

If T < 0 Or S_0 < 0 Or sigma < 0 Or K < 0 Then
   GoTo Exit_With_Error
End If

Dim d1 As Double: d1 = (Math.Log(S_0 / K) + (r + (sigma * sigma) / 2) * T) / (sigma * Sqr(T))
Dim d2 As Double: d2 = d1 - sigma * Sqr(T)

If IsCall = True Then
    Rho_Option = K * T * Exp(-r * T) * Application.WorksheetFunction.Norm_Dist(d2, 0, 1, True)
    Exit Function
Else
    Rho_Option = -K * T * Exp(-r * T) * Application.WorksheetFunction.Norm_Dist(-d2, 0, 1, True)
    Exit Function
End If


Exit_With_Error:
    Rho_Option = -99

End Function
