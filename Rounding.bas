Attribute VB_Name = "Rounding"
' RoundSignificant v1.0.1
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/CactusData/VBA.RoundSignificant
'
' Set of functions for rounding Currency, Decimal, and Double
' to a specified count of significant figures.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Compare Database
Option Explicit


' Common constants.
'
Public Const Base10     As Double = 10


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
