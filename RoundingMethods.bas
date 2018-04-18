Attribute VB_Name = "RoundingMethods"
' RoundingMethods v1.3.2
' (c) 2018-04-09. Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Round
'
' Set of functions for rounding Currency, Decimal, and Double
' up, down, by 4/5, to a specified count of significant figures,
' or as a sum.
'
' Set of functions for rounding Currency, Decimal, and Double
' up, down, or by 4/5 with the power of two and debugging, and
' for converting decimal numbers to integers and fractions.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Explicit

' Common constants.
'
' Base values.
Public Const Base2      As Double = 2
Public Const Base10     As Double = 10

' Enums.
'
Public Enum rmRoundingMethod
    Down = -1
    Midpoint = 0
    Up = 1
End Enum

' Returns Log 10 of Value.
'
' 2018-02-09. Gustav Brock, Cactus Data ApS, CPH.
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

' Returns Log 2 of Value.
'
' 2018-02-20. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Log2( _
    ByVal Value As Double) _
    As Double

    ' No error handling as this should be handled
    ' outside this function.
    '
    ' Example:
    '
    '     If MyValue > 0 then
    '         LogMyValue = Log2(MyValue)
    '     Else
    '         ' Do something else ...
    '     End If
    
    Log2 = Log(Value) / Log(Base2)

End Function

' Rounds Value down with count of decimals as specified with parameter NumDigitsAfterDecimal.
'
' Rounds to integer if NumDigitsAfterDecimal is zero.
'
' Optionally, rounds negative values towards zero.
'
' Uses CDec() to prevent bit errors of reals.
'
' Execution time is about 0.5탎 for rounding to integer,
' else about 1탎.
'
' 2018-02-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RoundDown( _
    ByVal Value As Variant, _
    Optional ByVal NumDigitsAfterDecimal As Long, _
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
        If NumDigitsAfterDecimal <> 0 Then
            Scaling = CDec(Base10 ^ NumDigitsAfterDecimal)
        Else
            Scaling = 1
        End If
        If Scaling = 0 Then
            ' A very large value for NumDigitsAfterDecimal has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        ElseIf RoundingToZero = False Then
            ' Round numeric value down.
            If Scaling = 1 Then
                ' Integer rounding.
                ReturnValue = Int(Value)
            Else
                ' First try with conversion to Decimal to avoid bit errors for some reals like 32.675.
                ' Very large values for NumDigitsAfterDecimal can cause an out-of-range error when dividing.
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
                ' Very large values for NumDigitsAfterDecimal can cause an out-of-range error when dividing.
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

' Rounds Value by 4/5 with count of decimals as specified with parameter NumDigitsAfterDecimal.
'
' Rounds to integer if NumDigitsAfterDecimal is zero.
'
' Rounds correctly Value until max/min value limited by a Scaling of 10
' raised to the power of (the number of decimals).
'
' Uses CDec() to prevent bit errors of reals.
'
' Execution time is about 1탎.
'
' 2018-02-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RoundMid( _
    ByVal Value As Variant, _
    Optional ByVal NumDigitsAfterDecimal As Long, _
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
        Scaling = CDec(Base10 ^ NumDigitsAfterDecimal)
        
        If Scaling = 0 Then
            ' A very large value for NumDigitsAfterDecimal has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        ElseIf MidwayRoundingToEven Then
            ' Banker's rounding.
            If Scaling = 1 Then
                ReturnValue = Round(Value)
            Else
                ' First try with conversion to Decimal to avoid bit errors for some reals like 32.675.
                ' Very large values for NumDigitsAfterDecimal can cause an out-of-range error when dividing.
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
            ' Very large values for NumDigitsAfterDecimal can cause an out-of-range error when dividing.
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

' Rounds Value up with count of decimals as specified with parameter NumDigitsAfterDecimal.
'
' Rounds to integer if NumDigitsAfterDecimal is zero.
'
' Optionally, rounds negative values away from zero.
'
' Uses CDec() to prevent bit errors of reals.
'
' Execution time is about 0.5탎 for rounding to integer,
' else about 1탎.
'
' 2018-02-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RoundUp( _
    ByVal Value As Variant, _
    Optional ByVal NumDigitsAfterDecimal As Long, _
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
        If NumDigitsAfterDecimal <> 0 Then
            Scaling = CDec(Base10 ^ NumDigitsAfterDecimal)
        Else
            Scaling = 1
        End If
        If Scaling = 0 Then
            ' A very large value for NumDigitsAfterDecimal has minimized scaling.
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
' Uses CDec() to prevent bit errors of reals.
'
' For rounding of values reaching the boundaries of type Currency, use the
' function RoundSignificantDec.
'
' Requires:
'   Function Log10.
'
' 2018-02-09. Gustav Brock, Cactus Data ApS, CPH.
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
' 2018-02-09. Gustav Brock, Cactus Data ApS, CPH.
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
' 2018-02-09. Gustav Brock, Cactus Data ApS, CPH.
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

' Rounds a series of numbers so the sum of these matches the
' rounded sum of the unrounded values.
' Further, if a requested total is passed, the rounded values
' will be scaled, so the sum of these matches the rounded total.
' In cases where the sum of the rounded values doesn't match
' the rounded total, the rounded values will be adjusted where
' the applied error will be the relatively smallest.
'
' The series of values to round must be passed as an array.
' The data type can be any numeric data type, and values can have
' any value.
' Internally, the function uses Decimal to achieve the highest
' precision and Double when the values exceed the range of Decimal.
'
' The result is an array holding the rounded values, as well as
' (by reference) the rounded total.
'
' If non-numeric values are passed, an error is raised.
'
' Requires:
'   RoundMid
'
' 2018-03-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RoundSum( _
    ByVal Values As Variant, _
    Optional ByRef Total As Variant, _
    Optional ByVal NumDigitsAfterDecimal As Long) _
    As Variant
    
    Dim SortedItems()   As Long
    Dim RoundedValues   As Variant
    Dim SortingValues   As Variant
    
    Dim Sum             As Variant
    Dim Value           As Variant
    Dim RoundedSum      As Variant
    Dim RoundedTotal    As Variant
    Dim PlusSum         As Variant
    Dim MinusSum        As Variant
    Dim RoundedPlusSum  As Variant
    Dim RoundedMinusSum As Variant
    
    Dim ErrorNumber     As Long
    Dim Item            As Long
    Dim Sign            As Variant
    Dim Ratio           As Variant
    Dim Difference      As Variant
    Dim Delta           As Variant
    Dim SortValue       As Variant
    
    ' Raise error if an array is not passed.
    Item = UBound(Values)
    
    ' Ignore errors while summing the values.
    On Error Resume Next
    If Err.Number = 0 Then
        ' Try to sum the passed values as a Decimal.
        Sum = CDec(0)
        For Item = LBound(Values) To UBound(Values)
            If IsNumeric(Values(Item)) Then
                Sum = Sum + CDec(Values(Item))
                If Err.Number <> 0 Then
                    ' Values exceed range of Decimal.
                    ' Exit loop and try using Double.
                    Exit For
                End If
            End If
        Next
    End If
    If Err.Number <> 0 Then
        ' Try to sum the passed values as a Double.
        Err.Clear
        Sum = CDbl(0)
        For Item = LBound(Values) To UBound(Values)
            If IsNumeric(Values(Item)) Then
                Sum = Sum + CDbl(Values(Item))
                If Err.Number <> 0 Then
                    ' Values exceed range of Double.
                    ' Exit loop and raise error.
                    Exit For
                End If
            End If
        Next
    End If
    ' Collect the error number as "On Error Goto 0" will clear it.
    ErrorNumber = Err.Number
    On Error GoTo 0
    If ErrorNumber <> 0 Then
        ' Extreme values. Give up.
        Error.Raise ErrorNumber
    End If
    
    ' Correct a missing or invalid parameter value for Total.
    If Not IsNumeric(Total) Then
        Total = 0
    End If
    If Total = 0 Then
        RoundedTotal = 0
    Else
        ' Round Total to an appropriate data type.
        ' Set data type of RoundedTotal to match Sum.
        Select Case VarType(Sum)
            Case vbSingle, vbDouble
                Value = CDbl(Total)
            Case Else
                Value = CDec(Total)
        End Select
        RoundedTotal = RoundMid(Value, NumDigitsAfterDecimal)
    End If
    
    ' Calculate scaling factor and sign.
    If Sum = 0 Or RoundedTotal = 0 Then
        ' Cannot scale a value of zero.
        Sign = 1
        Ratio = 1
    Else
        Sign = Sgn(Sum) * Sgn(RoundedTotal)
        ' Ignore error and convert to Double if total exceeds the range of Decimal.
        On Error Resume Next
        Ratio = Abs(RoundedTotal / Sum)
        If Err.Number <> 0 Then
            RoundedTotal = CDbl(RoundedTotal)
            Ratio = Abs(RoundedTotal / Sum)
        End If
        On Error GoTo 0
    End If
    
    ' Create array to hold the rounded values.
    RoundedValues = Values
    ' Scale and round the values and sum the rounded values.
    ' Variables will get the data type of RoundedValues.
    ' Ignore error and convert to Double if total exceeds the range of Decimal.
    On Error Resume Next
    For Item = LBound(Values) To UBound(Values)
        RoundedValues(Item) = RoundMid(Values(Item) * Ratio, NumDigitsAfterDecimal)
        If RoundedValues(Item) > 0 Then
            PlusSum = PlusSum + Values(Item)
            RoundedPlusSum = RoundedPlusSum + RoundedValues(Item)
            If Err.Number <> 0 Then
                RoundedPlusSum = CDbl(RoundedPlusSum) + CDbl(RoundedValues(Item))
            End If
        Else
            MinusSum = MinusSum + Values(Item)
            RoundedMinusSum = RoundedMinusSum + RoundedValues(Item)
            If Err.Number <> 0 Then
                RoundedMinusSum = CDbl(RoundedMinusSum) + CDbl(RoundedValues(Item))
            End If
        End If
    Next
    RoundedSum = RoundedPlusSum + RoundedMinusSum
    If Err.Number <> 0 Then
        RoundedPlusSum = CDbl(RoundedPlusSum)
        RoundedMinusSum = CDbl(RoundedMinusSum)
        RoundedSum = RoundedPlusSum + RoundedMinusSum
    End If
    On Error GoTo 0
    
    If RoundedTotal = 0 Then
        ' No total is requested.
        ' Use as total the rounded sum of the passed unrounded values.
        RoundedTotal = RoundMid(Sum, NumDigitsAfterDecimal)
    End If
    
    ' Check if a correction of the rounded values is needed.
    If (RoundedPlusSum + RoundedMinusSum = 0) And (RoundedTotal = 0) Then
        ' All items are rounded to zero. Nothing to do.
        ' Return zero.
    ElseIf RoundedSum = RoundedTotal Then
        ' Match. Nothing more to do.
    ElseIf RoundedSum = Sign * RoundedTotal Then
        ' Match, except that values shall be reversely signed.
        ' Will be done later before exit.
    Else
        ' Correction is needed.
        ' Redim array to hold the sorting of the rounded values.
        ReDim SortedItems(LBound(Values) To UBound(Values))
        ' Fill array with default sorting.
        For Item = LBound(SortedItems) To UBound(SortedItems)
            SortedItems(Item) = Item
        Next
        
        ' Create array to hold the values to sort.
        SortingValues = RoundedValues
        ' Fill the array after the relative rounding error and - for items with equal rounding error - the
        ' size of the value of items.
        For Item = LBound(SortedItems) To UBound(SortedItems)
            If Values(SortedItems(Item)) = 0 Then
                ' Zero value.
                SortValue = 0
            ElseIf RoundedPlusSum + RoundedMinusSum = 0 Then
                ' Values have been rounded to zero.
                ' Use original values.
                SortValue = Values(SortedItems(Item))
            ElseIf VarType(Values(SortedItems(Item))) = vbDouble Then
                ' Calculate relative rounding error.
                ' Value is exceeding Decimal. Use Double.
                SortValue = (Values(SortedItems(Item)) * Ratio - CDbl(RoundedValues(SortedItems(Item)))) * (Values(SortedItems(Item)) / Sum)
            Else
                ' Calculate relative rounding error using Decimal.
                SortValue = (Values(SortedItems(Item)) * Ratio - RoundedValues(SortedItems(Item))) * (Values(SortedItems(Item)) / Sum)
            End If
            ' Sort on the absolute value.
            SortingValues(Item) = Abs(SortValue)
        Next
        
        ' Sort the array after the relative rounding error and - for items with equal rounding error - the
        ' size of the value of items.
        QuickSortIndex SortedItems, SortingValues
        
        ' Distribute a difference between the rounded sum and the requested total.
        If RoundedPlusSum + RoundedMinusSum = 0 Then
            ' All rounded values are zero.
            ' Set Difference to the rounded total.
            Difference = RoundedTotal
        Else
            Difference = Sgn(RoundedSum) * (Abs(RoundedTotal) - Abs(RoundedSum))
        End If
        ' If Difference is positive, some values must be rounded up.
        ' If Difference is negative, some values must be rounded down.
        ' Calculate Delta, the value to increment/decrement by.
        Delta = Sgn(Difference) * 10 ^ -NumDigitsAfterDecimal
        
        ' Loop the rounded values and increment/decrement by Delta until Difference is zero.
        For Item = UBound(SortedItems) To LBound(SortedItems) Step -1
            ' If values should be incremented, ignore values rounded up.
            ' If values should be decremented, ignore values rounded down.
            If Sgn(Difference) = Sgn(Values(SortedItems(Item)) * Ratio - RoundedValues(SortedItems(Item))) Then
                ' Adjust this item.
                RoundedValues(SortedItems(Item)) = RoundedValues(SortedItems(Item)) + Delta
                If Item > LBound(SortedItems) Then
                    ' Check if the next item holds the exact reverse value.
                    If Values(SortedItems(Item)) = -Values(SortedItems(Item - 1)) Then
                        ' Adjust the next item as well to avoid uneven incrementing.
                        Item = Item - 1
                        RoundedValues(SortedItems(Item)) = RoundedValues(SortedItems(Item)) - Delta
                        Difference = Difference + Delta
                    End If
                End If
                Difference = Difference - Delta
            End If
            If Difference = 0 Then
                Exit For
            End If
        Next
    End If
    
    If Sign = -1 Then
        ' The values shall be reversely signed.
        For Item = LBound(RoundedValues) To UBound(RoundedValues)
            RoundedValues(Item) = -RoundedValues(Item)
        Next
    End If
    
    ' Return the rounded total.
    Total = RoundedTotal
    ' Return the array holding the rounded values.
    RoundSum = RoundedValues
    
End Function

' Rounds Value down to the power of two as specified with parameter Exponent.
'
' If Exponent is positive, the fraction of Value is rounded to an integer a fraction of 1 / 2 ^ Exponent.
' If Exponent is zero, Value is rounded to an integer.
' If Exponent is negative, Value is rounded to an integer and a multiplum of 2 ^ Exponent.
'
' Optionally, rounds negative values towards zero.
'
' Rounds correctly Value until max/min value limited by a scaling of 2 raised to the power of Exponent.
'
' Smallest fraction for rounding is:
'   2 ^ -21 (= 1 / 2097152)
' or:
'   0.000000476837158203125
'
' Largest numerical value to round with maximum resolution is:
'   79228162 + (2 ^ 21 - 8) / 2 ^ 21
' or:
'   79228162.999996185302734375
'
' Expected rounded value must not exceed the range of:
'   +/- 79,228,162,514,264,337,593,543,950,335
'
' Uses CDec() to prevent bit errors of reals.
'
' Execution time is about 0.5탎 for rounding to integer, else about 1탎.
'
' Examples, integers:
'   RoundDownBase2(1001, -3)                -> 1000
'   RoundDownBase2(1001, -8)                ->  768
'   RoundDownBase2(17.03, -4)               ->   16
'   RoundDownBase2(17.03, -5)               ->    0
'
' Examples, decimals:
'   1 / 2 ^ 4 = 0.0625                                  Step value when rounding by 1/16
'   RoundDownBase2(17.03, 4)                -> 17
'   RoundDownBase2(17.08, 4)                -> 17.0625
'   RoundDownBase2(17.1, 4)                 -> 17.0625
'   RoundDownBase2(17.2, 4)                 -> 17.1875
'
'   1 / 2 ^ 5 = 0.03125                                 Step value when rounding by 1/32
'   RoundDownBase2(17.125 + 0.00000, 4)     -> 17.125   Exact value. No rounding.
'   RoundDownBase2(17.125 + 0.03124, 4)     -> 17.125
'   RoundDownBase2(17.125 + 0.03125, 4)     -> 17.125
'
' More info on the power of two and rounding:
'   https://en.wikipedia.org/wiki/Power_of_two
'
' 2018-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RoundDownBase2( _
    ByVal Value As Variant, _
    Optional ByVal Exponent As Long, _
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
        If Exponent <> 0 Then
            Scaling = CDec(Base2 ^ Exponent)
        Else
            Scaling = 1
        End If
        If Scaling = 0 Then
            ' A very large value for Exponent has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        ElseIf RoundingToZero = False Then
            ' Round numeric value down.
            If Scaling = 1 Then
                ' Integer rounding.
                ReturnValue = Int(Value)
            Else
                ' First try with conversion to Decimal to avoid bit errors for some reals like 32.675.
                ' Very large values for Exponent can cause an out-of-range error when dividing.
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
                ' Very large values for NumDigitsAfterDecimal can cause an out-of-range error when dividing.
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
    
    RoundDownBase2 = ReturnValue

End Function

' Rounds Value by 4/5 to the power of two as specified with parameter Exponent.
'
' If Exponent is positive, the fraction of Value is rounded to an integer a fraction of 1 / 2 ^ Exponent.
' If Exponent is zero, Value is rounded to an integer.
' If Exponent is negative, Value is rounded to an integer and a multiplum of 2 ^ Exponent.
'
' Rounds correctly Value until max/min value limited by a scaling of 2 raised to the power of Exponent.
'
' Smallest fraction for rounding is:
'   2 ^ -21 (= 1 / 2097152)
' or:
'   0.000000476837158203125
'
' Largest numerical value to round with maximum resolution is:
'   79228162 + (2 ^ 21 - 8) / 2 ^ 21
' or:
'   79228162.999996185302734375
'
' Expected rounded value must not exceed the range of:
'   +/- 79,228,162,514,264,337,593,543,950,335
'
' Uses CDec() to prevent bit errors of reals.
'
' Execution time is about 1탎.
'
' Examples, integers:
'   RoundMidBase2(1001, -3)             -> 1000
'   RoundMidBase2(1001, -8)             -> 1024
'   RoundMidBase2(17.03, -4)            ->   16
'   RoundMidBase2(17.03, -5)            ->   32
'
' Examples, decimals:
'   1 / 2 ^ 4 = 0.0625                              Step value when rounding by 1/16
'   RoundMidBase2(17.03, 4)             -> 17.0     Rounding down
'   RoundMidBase2(17.08, 4)             -> 17.0625  Rounding down
'   RoundMidBase2(17.1, 4)              -> 17.125   Rounding up
'   RoundMidBase2(17.2, 4)              -> 17.1875  Rounding down
'
'   1 / 2 ^ 5 = 0.03125                             Step value when rounding by 1/32
'   RoundMidBase2(17.125 + 0.00000, 4)  -> 17.125   Exact value. No rounding.
'   RoundMidBase2(17.125 + 0.03124, 4)  -> 17.125   Rounding down
'   RoundMidBase2(17.125 + 0.03125, 4)  -> 17.1875  Rounding up
'
' More info on the power of two and rounding:
'   https://en.wikipedia.org/wiki/Power_of_two
'
' 2018-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RoundMidBase2( _
    ByVal Value As Variant, _
    Optional ByVal Exponent As Long) _
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
        Scaling = CDec(Base2 ^ Exponent)
        
        If Scaling = 0 Then
            ' A very large value for exponent has minimized scaling.
            ' Return Value as is.
            ReturnValue = Value
        Else
            ' Standard 4/5 rounding.
            ' Very large values for Exponent can cause an out-of-range error when dividing.
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
    
    RoundMidBase2 = ReturnValue

End Function

' Rounds Value up to the power of two as specified with parameter Exponent.
'
' If Exponent is positive, the fraction of Value is rounded to an integer a fraction of 1 / 2 ^ Exponent.
' If Exponent is zero, Value is rounded to an integer.
' If Exponent is negative, Value is rounded to an integer and a multiplum of 2 ^ Exponent.
'
' Optionally, rounds negative values away from zero.
'
' Rounds correctly Value until max/min value limited by a scaling of 2 raised to the power of Exponent.
'
' Smallest fraction for rounding is:
'   2 ^ -21 (= 1 / 2097152)
' or:
'   0.000000476837158203125
'
' Largest numerical value to round with maximum resolution is:
'   79228162 + (2 ^ 21 - 8) / 2 ^ 21
' or:
'   79228162.999996185302734375
'
' Expected rounded value must not exceed the range of:
'   +/- 79,228,162,514,264,337,593,543,950,335
'
' Uses CDec() to prevent bit errors of reals.
'
' Execution time is about 0.5탎 for rounding to integer, else about 1탎.
'
' Examples, integers:
'   RoundUpBase2(1001, -3)              -> 1008
'   RoundUpBase2(1001, -8)              -> 1024
'   RoundUpBase2(17.03, -4)             ->   32
'   RoundUpBase2(17.03, -5)             ->   32
'
' Examples, decimals:
'   1 / 2 ^ 4 = 0.0625                              Step value when rounding by 1/16
'   RoundUpBase2(17.03, 4)              -> 17.0625
'   RoundUpBase2(17.08, 4)              -> 17.125
'   RoundUpBase2(17.1, 4)               -> 17.125
'   RoundUpBase2(17.2, 4)               -> 17.25
'
'   1 / 2 ^ 5 = 0.03125                             Step value when rounding by 1/32
'   RoundUpBase2(17.125 + 0.00000, 4)   -> 17.125   Exact value. No rounding.
'   RoundUpBase2(17.125 + 0.03124, 4)   -> 17.1875
'   RoundUpBase2(17.125 + 0.03125, 4)   -> 17.1875
'
' More info on the power of two and rounding:
'   https://en.wikipedia.org/wiki/Power_of_two
'
' 2018-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RoundUpBase2( _
    ByVal Value As Variant, _
    Optional ByVal Exponent As Long, _
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
        If Exponent <> 0 Then
            Scaling = CDec(Base2 ^ Exponent)
        Else
            Scaling = 1
        End If
        If Scaling = 0 Then
            ' A very large value for Exponent has minimized scaling.
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
    
    RoundUpBase2 = ReturnValue

End Function

' Rounds and converts a decimal value to an integer and the fraction of an integer
' using 4/5 midpoint rounding, optionally rounding up or down.
'
' Rounding method is determined by parameter RoundingMethod.
' For rounding up or down, rounding of negative values can optionally be set to
' away-from-zero or towards-zero respectively by parameter RoundingAsAbsolute.
'
' Returns the rounded value as a decimal.
' Returns numerator and denominator of the fraction by reference.
'
' For general examples, see function RoundMidBase2, RoundUpBase2, and RoundDownBase2.
'
' Will, for example, convert decimal inches to integer inches and a fraction of inches.
' However, numerator and denominator of the fraction are returned by reference in the
' parameters Numerator and Denominator for the value to be formatted as text by the
' calling procedure.
'
' Example:
'   Value = 7.22
'   Exponent = 2    ' will round to 1/4.
'   Numerator = 0
'   Denominator = 0
'
'   Result = ConvertDecimalFractions(Value, Exponent, Numerator, Denominator)
'
'   Result = 7.25
'   Numerator = 1
'   Denominator = 4
'
'   Result = ConvertDecimalFractions(Value, Exponent, Numerator, Denominator, Up)
'
'   Result = 7.25
'   Numerator = 1
'   Denominator = 4
'
'   Result = ConvertDecimalFractions(Value, Exponent, Numerator, Denominator, Down)
'
'   Result = 7
'   Numerator = 0
'   Denominator = 0
'
' If negative, parameter Exponent determines the rounding of the fraction as
' 1 / 2 ^ Exponent with a maximum of 21 - or from 1 / 2 to 1 / 2097152.
' For inches, that is a range from 12.7 mm to about 12.1 nm.
'
' If zero or positive, parameter Exponent determines the rounding of the
' integer value with 2 ^ Exponent with a maximum of 21 - or from 1 to 2097152.
' For inches, that is a range from 25.4 mm to about 53.27 km.
'
' Also, se comments for the required functions:
'
'   RoundUpBase2
'   RoundMidBase2
'   RoundDownBase2
'
' 2018-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ConvertDecimalFractions( _
    ByVal Value As Variant, _
    ByVal Exponent As Integer, _
    Optional ByRef Numerator As Long, _
    Optional ByRef Denominator As Long, _
    Optional RoundingMethod As rmRoundingMethod = Midpoint, _
    Optional RoundingAsAbsolute As Boolean) _
    As Variant
    
    Dim Number      As Variant
    Dim Fraction    As Variant
    
    ' Validate rounding method.
    Select Case RoundingMethod
        Case Up, Midpoint, Down
            ' OK.
        Case Else
            ' Use default rounding method.
            RoundingMethod = Midpoint
    End Select
    
    If Exponent <= 0 Then
        ' Integer rounding only.
        Select Case RoundingMethod
            Case Up
                Number = RoundUpBase2(Value, Exponent, RoundingAsAbsolute)
            Case Midpoint
                Number = RoundMidBase2(Value, Exponent)
            Case Down
                Number = RoundDownBase2(Value, Exponent, RoundingAsAbsolute)
        End Select
        Fraction = 0
        Numerator = 0
        Denominator = 0
    Else
        ' Rounding with fractions.
        Number = Fix(CDec(Value))
        Select Case RoundingMethod
            Case Up
                Fraction = RoundUpBase2(Value - Number, Exponent, RoundingAsAbsolute)
            Case Midpoint
                Fraction = RoundMidBase2(Value - Number, Exponent)
            Case Down
                Fraction = RoundDownBase2(Value - Number, Exponent, RoundingAsAbsolute)
        End Select
        
        If Fraction = 0 Or Abs(Fraction) = 1 Then
            ' Fraction has been rounded to 0 or +/-1.
            Numerator = 0
            Denominator = 0
        Else
            ' Calculate numerator and denominator for the fraction.
            Denominator = Base2 ^ Exponent
            Numerator = Fraction * Denominator
            ' Find the smallest denominator.
            While Numerator Mod Base2 = 0
                Numerator = Numerator / Base2
                Denominator = Denominator / Base2
            Wend
        End If
    End If
    
    ConvertDecimalFractions = Number + Fraction
    
End Function

' Rounds a value and prints the result in its Base 2 components
' and as a decimal value, and in a integer-fraction format.
'
' Will accept values within +/- 2 ^ 96.
' Also, se comments for the required functions:
'
'   ConvertDecimalFractions
'   RoundMidBase2
'   Log2
'
' Example:
'
'   DebugBase2 746.873, 2
'   Exponent      2 ^ Exponent  Factor        Value         Fraction
'    9             512           1             512
'    8             256           0             0
'    7             128           1             128
'    6             64            1             64
'    5             32            1             32
'    4             16            0             0
'    3             8             1             8
'    2             4             0             0
'    1             2             1             2
'    0             1             0             0
'   Total:                                     746          3/4
'   Decimal:                                   746,75
'
' 2018-03-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub DebugBase2( _
    ByVal Value As Variant, _
    Optional ByVal Exponent As Long)
    
    ' Maximum possible exponent.
    Const MaxExponent2  As Long = 96

    Dim Exponent2       As Long
    Dim Number          As Variant
    Dim Rounded         As Variant
    Dim Factor          As Long
    Dim Sign            As Long
    Dim Numerator       As Long
    Dim Denominator     As Long

    If Not IsNumeric(Value) Then Exit Sub
    Sign = Sgn(Value)
    If Sign = 0 Then Exit Sub
    
    Number = CDec(Fix(Abs(Value)))
    
    ' Split and print the integer part.
    Debug.Print "Exponent", "2 ^ Exponent", "Factor", "Value", "Fraction"
    If Number > 0 Then
        ' Print each bit and value.
        For Exponent2 = Int(Log2(Number)) To 0 Step -1
            If Exponent2 = MaxExponent2 Then
                ' Cannot perform further calculation.
                Factor = 1
                Number = 0
            Else
                Factor = Int(Number / CDec(Base2 ^ Exponent2))
                Number = Number - CDec(Factor * Base2 ^ Exponent2)
            End If
            Debug.Print Exponent2, Base2 ^ Exponent2, Factor, Sign * Factor * Base2 ^ Exponent2
        Next
    Else
        ' Print zero values.
        Debug.Print 0, 0, 0, 0
    End If
    
    ' Find and print the fraction.
    Rounded = ConvertDecimalFractions(Value, Exponent, Numerator, Denominator)
    Debug.Print "Total:", , , CDec(Fix(Value)), Numerator & "/" & Denominator
    Debug.Print "Decimal:", , , Rounded
    
End Sub

