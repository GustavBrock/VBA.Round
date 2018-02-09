Attribute VB_Name = "RoundingMethods"
' RoundingMethods v1.2.0
' (c) 2018-02-09. Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Round
'
' Set of functions for rounding Currency, Decimal, and Double
' up, down, by 4/5, or to a specified count of significant figures.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Explicit

' Common constants.
'
Public Const Base10     As Double = 10

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
            ' A very large value for Digits has minimized scaling.
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
' The data type can be any, and values can have any value.
' Internally, the function uses Decimal to achieve the highest
' precision and Double when the values exceed the range of Decimal.
'
' Result is an array holding the rounded values, as well as
' (by reference) the rounded total.
'
' If non-numeric values are passed, an error is raised.
'
' Requires:
'   RoundMid
'
' 2018-02-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RoundSum( _
    ByVal Values As Variant, _
    Optional ByRef Total As Variant, _
    Optional ByVal NumDigitsAfterDecimal As Long) _
    As Variant
    
    Dim SortedItems()   As Long
    Dim RoundedValues   As Variant
    
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
    Dim SortItem        As Long
    Dim ThisItem        As Long
    Dim SortRelation    As Variant
    Dim ThisRelation    As Variant
    Dim Sign            As Variant
    Dim Ratio           As Variant
    Dim Difference      As Variant
    Dim Delta           As Variant
    
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
        RoundedTotal = RoundMid(PlusSum + MinusSum, NumDigitsAfterDecimal)
    End If
    
    ' Check if a correction of the rounded values is needed.
    If RoundedPlusSum - RoundedMinusSum = 0 Then
        ' All items are zero. Nothing to do.
        ' Return zero.
        RoundedTotal = 0
    ElseIf RoundedSum = RoundedTotal Then
        ' Match. Nothing more to do.
    ElseIf RoundedSum = Sign * RoundedTotal Then
        ' Match, except that values shall be reversely signed.
        ' Will be done later before exit.
    Else
        ' Correction is needed.
        ' Create array to hold the sorting of the rounded values.
        ReDim SortedItems(LBound(Values) To UBound(Values))
        ' Fill array with default sorting.
        For Item = LBound(SortedItems) To UBound(SortedItems)
            SortedItems(Item) = Item
        Next
        ' Sort the array after the rounding error and - for items with equal rounding error - the
        ' size of the value of items.
        For Item = LBound(SortedItems) To UBound(SortedItems) - 1
            If Values(SortedItems(Item)) = 0 Then
                ThisRelation = 0
            ElseIf VarType(Values(SortedItems(Item))) = vbDouble Then
                ' Value is exceeding Decimal. Use Double.
                ThisRelation = (Values(SortedItems(Item)) * Ratio - CDbl(RoundedValues(SortedItems(Item)))) / Values(SortedItems(Item))
            Else
                ThisRelation = (Values(SortedItems(Item)) * Ratio - RoundedValues(SortedItems(Item))) / Values(SortedItems(Item))
            End If
            For SortItem = Item + 1 To UBound(SortedItems)
                If Values(SortedItems(SortItem)) = 0 Then
                    SortRelation = 0
                ElseIf VarType(Values(SortedItems(SortItem))) = vbDouble Then
                    ' Value is exceeding Decimal. Use Double.
                    SortRelation = (Values(SortedItems(SortItem)) * Ratio - CDbl(RoundedValues(SortedItems(SortItem)))) / Values(SortedItems(SortItem))
                Else
                    SortRelation = (Values(SortedItems(SortItem)) * Ratio - RoundedValues(SortedItems(SortItem))) / Values(SortedItems(SortItem))
                End If
                If Abs(ThisRelation) >= Abs(SortRelation) Or (ThisRelation = SortRelation And Abs(RoundedValues(SortedItems(Item))) >= Abs(RoundedValues(SortedItems(SortItem)))) Then
                    ThisItem = SortedItems(Item)
                    SortedItems(Item) = SortedItems(SortItem)
                    SortedItems(SortItem) = ThisItem
                End If
            Next
        Next

        ' Distribute a difference between the rounded sum and the requested total.
        Difference = Sgn(RoundedSum) * (Abs(RoundedTotal) - Abs(RoundedSum))
        ' If Difference is positive, some values must be rounded up.
        ' If Difference is negative, some values must be rounded down.
        ' Calculate Delta, the value to increment/decrement by.
        Delta = Sgn(Difference) * 10 ^ -NumDigitsAfterDecimal
        
        ' Loop the rounded values and increment/decrement by Delta until Difference is zero.
        For Item = UBound(SortedItems) To LBound(SortedItems) Step -1
            If Sgn(Difference) = Sgn(Values(SortedItems(Item)) * Ratio - RoundedValues(SortedItems(Item))) Then
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
