Attribute VB_Name = "RoundingSignificantTest"
' RoundSignificantTest v1.0.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.RoundSignificant
'
' Test function to list rounding of example values.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Compare Database
Option Explicit

Public Function RoundingSignificantDemo()

    Dim Value               As Variant
    Dim RoundToEven         As Variant
    Dim RoundAwayFromZero   As Variant
    Dim Digits              As Integer
    Dim I                   As Integer
    
    Value = CCur(30.675)
    Digits = 14
    
    Debug.Print "Cur 14"
    For I = 0 To 10
        RoundToEven = RoundSignificantCur(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantCur(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CCur(30.675)
    Digits = 4
    
    Debug.Print "Cur 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantCur(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantCur(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    
    Value = CDec(30.675)
    Digits = 28
    
    Debug.Print "Dec 28"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDec(30.675)
    Digits = 4
    
    Debug.Print "Dec 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    
    Value = CDec(-30.675)
    Digits = 28
    
    Debug.Print "Dec 28"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDec(-30.675)
    Digits = 4
    
    Debug.Print "Dec 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    
    Value = CDec(-30.675) * 10 ^ 9
    Digits = 28
    
    Debug.Print "Dec 28"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I * 10 ^ 9 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I * 10 ^ 9 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDec(-30.675) * 10 ^ 9
    Digits = 4
    
    Debug.Print "Dec 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I * 10 ^ 9 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I * 10 ^ 9 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    
    Value = CDbl(30.675)
    Digits = 14
    
    Debug.Print "Dbl 14"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDbl(30.675)
    Digits = 4
    
    Debug.Print "Dbl 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print

    Value = CDbl(30.675) * 10 ^ 300
    Digits = 14
    
    Debug.Print "Dbl 14"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I * 10 ^ 300 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I * 10 ^ 300 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDbl(30.675) * 10 ^ 300
    Digits = 4
    
    Debug.Print "Dbl 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I * 10 ^ 300 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I * 10 ^ 300 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print

    Value = CDbl(30.675) * 10 ^ -300
    Digits = 8
    
    Debug.Print "Dbl 8"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I * 10 ^ -300 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I * 10 ^ -300 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDbl(30.675) * 10 ^ -300
    Digits = 4
    
    Debug.Print "Dbl 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I * 10 ^ -300 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I * 10 ^ -300 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print

End Function
