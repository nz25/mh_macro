Attribute VB_Name = "Stats"
Option Explicit

Public Function GetProbability(n1 As Double, x1 As Double, n2 As Double, x2 As Double) As Double

    'n1 - base 1, x1 - value 1, n2 = base 2, x2 - value 2
    'x1 and x2 values should be between 0 and 1, for example 45%
    If x1 < 0 Or x1 > 1 Or x2 < 0 Or x2 > 1 Then
        Err.Raise Globals.GetProbabilityParameterErrorCode, "Error in Statistics module", "Parameter validation error at Statistics.GetPValue"
    End If
    
    'Special cases resulting in errors -> should be NOT significant
    If x1 + x2 = 0 Or x1 = x2 Or n1 < Globals.BaseTooLow Or n2 < Globals.BaseTooLow Then  'not significant
        GetProbability = 1
        Exit Function
    End If

    'A special case resulting in division by zero -> should be significcant
    If (x1 = 1 And x2 = 0) Or (x1 = 0 And x2 = 1) Then
        GetProbability = 0
        Exit Function
    End If
    
    
    Dim a As Double, b As Double, c As Double, d As Double, f As Double, x As Double

    a = x2 - x1
    b = Sqr((n2 * n1) / (n2 + n1))
    c = n1 * x1 * (1 - x1)
    d = n2 * x2 * (1 - x2)
    f = n2 + n1 - 2
    x = Abs(a * b / Sqr((c + d) / f))
    
    GetProbability = WorksheetFunction.T_Dist_2T(x, f)

End Function

Public Function IsSignificant(SignificanceLevel As Double, probability As Double) As Boolean

    'significanceLevel should be between 0 and 1. E.g. for a test with 95% confidence = 0.95
    If SignificanceLevel < 0 Or SignificanceLevel > 1 Then
        Err.Raise Globals.IsSignficantErrorCode, "Error in Statistics module", "Parameter validation error at Statistics.IsSignificant"
    End If

    If probability <= 1 - SignificanceLevel Then
        IsSignificant = True
    Else
        IsSignificant = False
    End If
    
End Function

