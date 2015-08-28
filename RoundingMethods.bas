Attribute VB_Name = "RoundingMethods"
' RoundingMethods v1.0.1
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Round
'
' Set of functions for rounding Currency, Decimal, and Double
' up, down, by 4/5, or to a specified count of significant figures.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Compare Database
Option Explicit

' Common constants.
'
Public Const Base10     As Double = 10

' Rounds Value by 4/5 with count of decimals as specified with parameter NumDigitsAfterDecimals.
'
' Rounds to integer if NumDigitsAfterDecimals is zero.
'
' Rounds correctly Value until max/min value limited by a Scaling of 10
' raised to the power of (the number of decimals).
'
' Uses CDec() for correcting bit errors of reals.
'
' Execution time is about 1µs.
'
Public Function RoundMid( _
    ByVal Value As Variant, _
    Optional ByVal NumDigitsAfterDecimals As Long, _
    Optional ByVal MidwayRoundingToEven As Boolean) _
    As Variant

    Dim Scaling     As Variant
    Dim Half        As Variant
    Dim ScaledValue As Variant
    Dim ReturnValue As Variant
    
    ' Only round if Value is numeric and ReturnValue can be different from zero.
    If Not IsNumeric(Value) Then
        ' Nothing to do.
        ReturnValue = Null
    ElseIf Value = 0 Then
        ' Nothing to round.
        ' Return Value as is.
        ReturnValue = Value
    Else
        Scaling = CDec(Base10 ^ NumDigitsAfterDecimals)
        
        If Scaling = 0 Then
            ' A very large value for Digits has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        ElseIf MidwayRoundingToEven Then
            ' Banker's rounding.
            If Scaling = 1 Then
                ReturnValue = Round(Value)
            Else
                ' First try with conversion to Decimal to avoid bit errors for some reals like 32.675.
                ' Very large values for NumDigitsAfterDecimals can cause an out-of-range error when dividing.
                On Error Resume Next
                ScaledValue = Round(CDec(Value) * Scaling)
                ReturnValue = ScaledValue / Scaling
                If Err.Number <> 0 Then
                    ' Decimal overflow.
                    ' Round Value without conversion to Decimal.
                    ReturnValue = Round(Value * Scaling) / Scaling
                End If
            End If
        Else
            ' Standard 4/5 rounding.
            ' Very large values for NumDigitsAfterDecimals can cause an out-of-range error when dividing.
            On Error Resume Next
            Half = CDec(0.5)
            If Value > 0 Then
                ScaledValue = Int(CDec(Value) * Scaling + Half)
            Else
                ScaledValue = -Int(-CDec(Value) * Scaling + Half)
            End If
            ReturnValue = ScaledValue / Scaling
            If Err.Number <> 0 Then
                ' Decimal overflow.
                ' Round Value without conversion to Decimal.
                Half = CDbl(0.5)
                If Value > 0 Then
                    ScaledValue = Int(Value * Scaling + Half)
                Else
                    ScaledValue = -Int(-Value * Scaling + Half)
                End If
                ReturnValue = ScaledValue / Scaling
            End If
        End If
        If Err.Number <> 0 Then
            ' Rounding failed because values are near one of the boundaries of type Double.
            ' Return value as is.
            ReturnValue = Value
        End If
    End If
    
    RoundMid = ReturnValue

End Function

' Rounds Value down with count of decimals as specified with parameter NumDigitsAfterDecimals.
'
' Rounds to integer if NumDigitsAfterDecimals is zero.
'
' Optionally, rounds negative values towards zero.
'
' Uses CDec() for correcting bit errors of reals.
'
' Execution time is about 0.5µs for rounding to integer
' else about 1µs.
'
Public Function RoundDown( _
    ByVal Value As Variant, _
    Optional ByVal NumDigitsAfterDecimals As Long, _
    Optional ByVal RoundingToZero As Boolean) _
    As Variant
    
    Dim Scaling     As Variant
    Dim ScaledValue As Variant
    Dim ReturnValue As Variant
    
    ' Only round if Value is numeric and ReturnValue can be different from zero.
    If Not IsNumeric(Value) Then
        ' Nothing to do.
        ReturnValue = Null
    ElseIf Value = 0 Then
        ' Nothing to round.
        ' Return Value as is.
        ReturnValue = Value
    Else
        If NumDigitsAfterDecimals <> 0 Then
            Scaling = CDec(Base10 ^ NumDigitsAfterDecimals)
        Else
            Scaling = 1
        End If
        If Scaling = 0 Then
            ' A very large value for Digits has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        ElseIf RoundingToZero = False Then
            ' Round numeric value down.
            If Scaling = 1 Then
                ' Integer rounding.
                ReturnValue = Int(Value)
            Else
                ' First try with conversion to Decimal to avoid bit errors for some reals like 32.675.
                ' Very large values for NumDigitsAfterDecimals can cause an out-of-range error when dividing.
                On Error Resume Next
                ScaledValue = Int(CDec(Value) * Scaling)
                ReturnValue = ScaledValue / Scaling
                If Err.Number <> 0 Then
                    ' Decimal overflow.
                    ' Round Value without conversion to Decimal.
                    ScaledValue = Int(Value * Scaling)
                    ReturnValue = ScaledValue / Scaling
                End If
            End If
        Else
            ' Round absolute value down.
            If Scaling = 1 Then
                ' Integer rounding.
                ReturnValue = Fix(Value)
            Else
                ' First try with conversion to Decimal to avoid bit errors for some reals like 32.675.
                ' Very large values for NumDigitsAfterDecimals can cause an out-of-range error when dividing.
                On Error Resume Next
                ScaledValue = Fix(CDec(Value) * Scaling)
                ReturnValue = ScaledValue / Scaling
                If Err.Number <> 0 Then
                    ' Decimal overflow.
                    ' Round Value with no conversion.
                    ScaledValue = Fix(Value * Scaling)
                    ReturnValue = ScaledValue / Scaling
                End If
            End If
        End If
        If Err.Number <> 0 Then
            ' Rounding failed because values are near one of the boundaries of type Double.
            ' Return value as is.
            ReturnValue = Value
        End If
    End If
    
    RoundDown = ReturnValue

End Function

' Rounds Value up with count of decimals as specified with parameter NumDigitsAfterDecimals.
'
' Rounds to integer if NumDigitsAfterDecimals is zero.
'
' Optionally, rounds negative values away from zero.
'
' Uses CDec() for correcting bit errors of reals.
'
' Execution time is about 0.5µs for rounding to integer
' else about 1µs.
'
Public Function RoundUp( _
    ByVal Value As Variant, _
    Optional ByVal NumDigitsAfterDecimals As Long, _
    Optional ByVal RoundingAwayFromZero As Boolean) _
    As Variant

    Dim Scaling     As Variant
    Dim ScaledValue As Variant
    Dim ReturnValue As Variant
    
    ' Only round if Value is numeric and ReturnValue can be different from zero.
    If Not IsNumeric(Value) Then
        ' Nothing to do.
        ReturnValue = Null
    ElseIf Value = 0 Then
        ' Nothing to round.
        ' Return Value as is.
        ReturnValue = Value
    Else
        If NumDigitsAfterDecimals <> 0 Then
            Scaling = CDec(Base10 ^ NumDigitsAfterDecimals)
        Else
            Scaling = 1
        End If
        If Scaling = 0 Then
            ' A very large value for Digits has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        ElseIf RoundingAwayFromZero = False Or Value > 0 Then
            ' Round numeric value up.
            If Scaling = 1 Then
                ' Integer rounding.
                ReturnValue = -Int(-Value)
            Else
                ' First try with conversion to Decimal to avoid bit errors for some reals like 32.675.
                On Error Resume Next
                ScaledValue = -Int(CDec(-Value) * Scaling)
                ReturnValue = ScaledValue / Scaling
                If Err.Number <> 0 Then
                    ' Decimal overflow.
                    ' Round Value without conversion to Decimal.
                    ScaledValue = -Int(-Value * Scaling)
                    ReturnValue = ScaledValue / Scaling
                End If
            End If
        Else
            ' Round absolute value up.
            If Scaling = 1 Then
                ' Integer rounding.
                ReturnValue = Int(Value)
            Else
                ' First try with conversion to Decimal to avoid bit errors for some reals like 32.675.
                On Error Resume Next
                ScaledValue = Int(CDec(Value) * Scaling)
                ReturnValue = ScaledValue / Scaling
                If Err.Number <> 0 Then
                    ' Decimal overflow.
                    ' Round Value without conversion to Decimal.
                    ScaledValue = Int(Value * Scaling)
                    ReturnValue = ScaledValue / Scaling
                End If
            End If
        End If
        If Err.Number <> 0 Then
            ' Rounding failed because values are near one of the boundaries of type Double.
            ' Return value as is.
            ReturnValue = Value
        End If
    End If
    
    RoundUp = ReturnValue

End Function

' Rounds Value to have significant figures as specified with parameter Digits.
'
' Performs no rounding if Digits is zero.
' Rounds to integer if NoDecimals is True.
'
' Rounds correctly Value until max/min value of currency type multiplied with 10
' raised to the power of (the number of digits of the index of Value) minus Digits.
' This equals roughly +/-922 * 10 ^ 12 for any Value of Digits.
'
' Uses CDec() for correcting bit errors of reals.
'
' For rounding of values reaching the boundaries of type Currency, use the
' function RoundSignificantDec.
'
' Requires:
'   Function Log10.
'
Public Function RoundSignificantCur( _
    ByVal Value As Currency, _
    ByVal Digits As Integer, _
    Optional ByVal NoDecimals As Boolean, _
    Optional ByVal MidwayRoundingToEven As Boolean) _
    As Variant
    
    Dim Exponent    As Double
    Dim Scaling     As Double
    Dim Half        As Variant
    Dim ScaledValue As Variant
    Dim ReturnValue As Currency
    
    ' Only round if Value is numeric and result can be different from zero.
    If Not IsNumeric(Value) Then
        ' Nothing to do.
        ReturnValue = Null
    ElseIf (Value = 0 Or Digits <= 0) Then
        ' Nothing to round.
        ' Return Value as is.
        ReturnValue = Value
    Else
        ' Calculate scaling factor.
        Exponent = Int(Log10(Abs(Value))) + 1 - Digits
        If NoDecimals = True Then
            ' No decimals.
            If Exponent < 0 Then
                Exponent = 0
            End If
        End If
        Scaling = Base10 ^ Exponent
        
        If Scaling = 0 Then
            ' A very large value for Digits has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        Else
            ' Very large values for Digits can cause an out-of-range error when dividing.
            On Error Resume Next
            ScaledValue = CDec(Value) / Scaling
            If Err.Number <> 0 Then
                ' Return value as is.
                ReturnValue = Value
            Else
                ' Perform rounding.
                If MidwayRoundingToEven = False Then
                    ' Round away from zero.
                    Half = CDec(Sgn(Value) / 2)
                    ReturnValue = CCur(Fix(ScaledValue + Half) * Scaling)
                Else
                    ' Round to even.
                    ReturnValue = CCur(Round(ScaledValue) * Scaling)
                End If
                If Err.Number <> 0 Then
                    ' Rounding failed because values are near one of the boundaries of type Currency.
                    ' Return value as is.
                    ReturnValue = Value
                End If
            End If
        End If
    End If
  
    RoundSignificantCur = ReturnValue

End Function

' Rounds Value to have significant figures as specified with parameter Digits.
'
' Performs no rounding if Digits is zero.
' Rounds to integer if NoDecimals is True.
'
' Rounds correctly values until about max/min Value of Decimal type divided by 2.
' This equals roughly +/-4 * 10^28.
' Digits can be any value between 1 and 28.
' Also rounds correctly values less than numeric 1 with up to 28 decimals.
' Digits can then be any value between 1 and 27.
'
' Requires:
'   Function Log10.
'
Public Function RoundSignificantDec( _
    ByVal Value As Variant, _
    ByVal Digits As Integer, _
    Optional ByVal NoDecimals As Boolean, _
    Optional ByVal MidwayRoundingToEven As Boolean) _
    As Variant
    
    Dim Exponent    As Double
    Dim Scaling     As Variant
    Dim Half        As Variant
    Dim ScaledValue As Variant
    Dim ReturnValue As Variant
    
    ' Only round if Value is numeric and result can be different from zero.
    If Not IsNumeric(Value) Then
        ' Nothing to do.
        ReturnValue = Null
    ElseIf (Value = 0 Or Digits <= 0) Then
        ' Nothing to round.
        ' Return Value as is.
        ReturnValue = Value
    Else
        ' Calculate scaling factor.
        Exponent = Int(Log10(Abs(Value))) + 1 - Digits
        If NoDecimals = True Then
            ' No decimals.
            If Exponent < 0 Then
                Exponent = 0
            End If
        End If
        Scaling = CDec(Base10 ^ Exponent)
        
        If Scaling = 0 Then
            ' A very large value for Digits has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        Else
            ' Very large values for Digits can cause an out-of-range error when dividing.
            On Error Resume Next
            ScaledValue = CDec(Value) / Scaling
            If Err.Number <> 0 Then
                ' Return value as is.
                ReturnValue = Value
            Else
                ' Perform rounding.
                If MidwayRoundingToEven = False Then
                    ' Round away from zero.
                    Half = CDec(Sgn(Value) / 2)
                    ReturnValue = Fix(ScaledValue + Half) * Scaling
                Else
                    ' Round to even.
                    ReturnValue = Round(ScaledValue) * Scaling
                End If
                If Err.Number <> 0 Then
                    ' Rounding failed because values are near one of the boundaries of type Decimal.
                    ' Return value as is.
                    ReturnValue = Value
                End If
            End If
        End If
    End If
  
    RoundSignificantDec = ReturnValue

End Function

' Rounds Value to have significant figures as specified with parameter Digits.
'
' Performs no rounding if Digits is zero.
' Rounds to integer if NoDecimals is True.
' Digits can be any value between 1 and 14.
'
' Will accept values until about max/min Value of Double type.
' At extreme values (beyond approx. E+/-300) with significant
' figures of 10 and above, rounding is not 100% perfect due to
' the limited precision of Double.
'
' For rounding of values within the range of type Decimal, use the
' function RoundSignificantDec.
'
' Requires:
'   Function Log10.
'
Public Function RoundSignificantDbl( _
    ByVal Value As Double, _
    ByVal Digits As Integer, _
    Optional ByVal NoDecimals As Boolean, _
    Optional ByVal MidwayRoundingToEven As Boolean) _
    As Double
    
    Dim Exponent    As Double
    Dim Scaling     As Double
    Dim Half        As Variant
    Dim ScaledValue As Variant
    Dim ReturnValue As Double
    
    ' Only round if result can be different from zero.
    If (Value = 0 Or Digits <= 0) Then
        ' Nothing to round.
        ' Return Value as is.
        ReturnValue = Value
    Else
        ' Calculate scaling factor.
        Exponent = Int(Log10(Abs(Value))) + 1 - Digits
        If NoDecimals = True Then
            ' No decimals.
            If Exponent < 0 Then
                Exponent = 0
            End If
        End If
        Scaling = Base10 ^ Exponent
        
        If Scaling = 0 Then
            ' A very large value for Digits has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        Else
            ' Very large values for Digits can cause an out-of-range error when dividing.
            On Error Resume Next
            ScaledValue = CDec(Value / Scaling)
            If Err.Number <> 0 Then
                ' Return value as is.
                ReturnValue = Value
            Else
                ' Perform rounding.
                If MidwayRoundingToEven = False Then
                    ' Round away from zero.
                    Half = CDec(Sgn(Value) / 2)
                    ReturnValue = CDbl(Fix(ScaledValue + Half)) * Scaling
                Else
                    ' Round to even.
                    ReturnValue = CDbl(Round(ScaledValue)) * Scaling
                End If
                If Err.Number <> 0 Then
                    ' Rounding failed because values are near one of the boundaries of type Double.
                    ' Return value as is.
                    ReturnValue = Value
                End If
            End If
        End If
    End If
  
    RoundSignificantDbl = ReturnValue

End Function

' Returns Log 10 of Value.
'
Public Function Log10( _
    ByVal Value As Double) _
    As Double

    ' No error handling as this should be handled
    ' outside this function.
    '
    ' Example:
    '
    '     If MyValue > 0 then
    '         LogMyValue = Log10(MyValue)
    '     Else
    '         ' Do something else ...
    '     End If
    
    Log10 = Log(Value) / Log(Base10)

End Function

