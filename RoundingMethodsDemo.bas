Attribute VB_Name = "RoundingMethodsDemo"
' RoundingMethodsDemo v1.3.0
' (c) 2024-03-11. Gustav Brock, Cactus Data ApS, CPH.
' https://github.com/GustavBrock/VBA.Round
'
' Demo functions to:
'   - list rounding of example values
'   - round a field of a recordset
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Explicit

' Create data for the rounding table in the EE article.
'
' 2018-02-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Demo()

    Dim n(0 To 7) As Double
    
    n(0) = 12.344
    n(1) = 12.345
    n(2) = 12.346
    n(3) = 12.354
    n(4) = 12.355
    n(5) = 12.356
    
    Debug.Print "Value n", , , n(0), n(1), n(2), n(3), n(4), n(5)
    Debug.Print "RoundUp(n, 2, False)", , RoundUp(n(0), 2, False), RoundUp(n(1), 2, False), RoundUp(n(2), 2, False), RoundUp(n(3), 2, False), RoundUp(n(4), 2, False), RoundUp(n(5), 2, False)
    Debug.Print "RoundUp(n, 2, True)", , RoundUp(n(0), 2, True), RoundUp(n(1), 2, True), RoundUp(n(2), 2, True), RoundUp(n(3), 2, True), RoundUp(n(4), 2, True), RoundUp(n(5), 2, True)
    Debug.Print "RoundDown(n, 2, False)", , RoundDown(n(0), 2, False), RoundDown(n(1), 2, False), RoundDown(n(2), 2, False), RoundDown(n(3), 2, False), RoundDown(n(4), 2, False), RoundDown(n(5), 2, False)
    Debug.Print "RoundDown(n, 2, True)", , RoundDown(n(0), 2, True), RoundDown(n(1), 2, True), RoundDown(n(2), 2, True), RoundDown(n(3), 2, True), RoundDown(n(4), 2, True), RoundDown(n(5), 2, True)
    Debug.Print "RoundMid(n, 2, False)", , RoundMid(n(0), 2, False), RoundMid(n(1), 2, False), RoundMid(n(2), 2, False), RoundMid(n(3), 2, False), RoundMid(n(4), 2, False), RoundMid(n(5), 2, False)
    Debug.Print "RoundMid(n, 2, True)", , RoundMid(n(0), 2, True), RoundMid(n(1), 2, True), RoundMid(n(2), 2, True), RoundMid(n(3), 2, True), RoundMid(n(4), 2, True), RoundMid(n(5), 2, True)
    Debug.Print "RoundSignificantDec(n, 4, , False)", RoundSignificantDec(n(0), 4, , False), RoundSignificantDec(n(1), 4, , False), RoundSignificantDec(n(2), 4, , False), RoundSignificantDec(n(3), 4, , False), RoundSignificantDec(n(4), 4, , False), RoundSignificantDec(n(5), 4, , False)
    Debug.Print "RoundSignificantDec(n, 4, , True)", RoundSignificantDec(n(0), 4, , True), RoundSignificantDec(n(1), 4, , True), RoundSignificantDec(n(2), 4, , True), RoundSignificantDec(n(3), 4, , True), RoundSignificantDec(n(4), 4, , True), RoundSignificantDec(n(5), 4, , True)
    
    n(0) = -n(0)
    n(1) = -n(1)
    n(2) = -n(2)
    n(3) = -n(3)
    n(4) = -n(4)
    n(5) = -n(5)
    
    Debug.Print
    Debug.Print "Value n", , , n(0), n(1), n(2), n(3), n(4), n(5)
    Debug.Print "RoundUp(n, 2, False)", , RoundUp(n(0), 2, False), RoundUp(n(1), 2, False), RoundUp(n(2), 2, False), RoundUp(n(3), 2, False), RoundUp(n(4), 2, False), RoundUp(n(5), 2, False)
    Debug.Print "RoundUp(n, 2, True)", , RoundUp(n(0), 2, True), RoundUp(n(1), 2, True), RoundUp(n(2), 2, True), RoundUp(n(3), 2, True), RoundUp(n(4), 2, True), RoundUp(n(5), 2, True)
    Debug.Print "RoundDown(n, 2, False)", , RoundDown(n(0), 2, False), RoundDown(n(1), 2, False), RoundDown(n(2), 2, False), RoundDown(n(3), 2, False), RoundDown(n(4), 2, False), RoundDown(n(5), 2, False)
    Debug.Print "RoundDown(n, 2, True)", , RoundDown(n(0), 2, True), RoundDown(n(1), 2, True), RoundDown(n(2), 2, True), RoundDown(n(3), 2, True), RoundDown(n(4), 2, True), RoundDown(n(5), 2, True)
    Debug.Print "RoundMid(n, 2, False)", , RoundMid(n(0), 2, False), RoundMid(n(1), 2, False), RoundMid(n(2), 2, False), RoundMid(n(3), 2, False), RoundMid(n(4), 2, False), RoundMid(n(5), 2, False)
    Debug.Print "RoundMid(n, 2, True)", , RoundMid(n(0), 2, True), RoundMid(n(1), 2, True), RoundMid(n(2), 2, True), RoundMid(n(3), 2, True), RoundMid(n(4), 2, True), RoundMid(n(5), 2, True)
    Debug.Print "RoundSignificantDec(n, 4, , False)", RoundSignificantDec(n(0), 4, , False), RoundSignificantDec(n(1), 4, , False), RoundSignificantDec(n(2), 4, , False), RoundSignificantDec(n(3), 4, , False), RoundSignificantDec(n(4), 4, , False), RoundSignificantDec(n(5), 4, , False)
    Debug.Print "RoundSignificantDec(n, 4, , True)", RoundSignificantDec(n(0), 4, , True), RoundSignificantDec(n(1), 4, , True), RoundSignificantDec(n(2), 4, , True), RoundSignificantDec(n(3), 4, , True), RoundSignificantDec(n(4), 4, , True), RoundSignificantDec(n(5), 4, , True)

End Function

' Create data for the cake split graph in the EE article.
'
' 2018-03-16. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub CheeseCakeSplit()

    Const Parts         As Integer = 12
    
    Dim Players(0 To 6) As Double
    Dim Shares          As Variant
    
    Dim Points          As Double
    Dim Player          As Integer
    Dim SumSlices       As Double
    Dim RequestedSlices As Integer

    Players(0) = 33
    Players(1) = 9
    Players(2) = 13
    Players(3) = 22
    Players(4) = 41
    Players(5) = 11
    Players(6) = 23
    
    For Player = LBound(Players) To UBound(Players)
        Points = Points + Players(Player)
    Next
    
    Shares = RoundSum(Players, Parts, 0)
    Debug.Print "Player", "Points", "Share", "Slices", "Rounded", "Error", "Corrected", "Result"
    For Player = LBound(Players) To UBound(Players)
        SumSlices = SumSlices + CDbl(Format(Parts * Players(Player) / Points, "0.000"))
        RequestedSlices = RequestedSlices + RoundMid(Parts * Players(Player) / Points, 0)
        Debug.Print _
            Player, _
            Players(Player), _
            Format(Players(Player) / Points, "0.0000"), _
            Format(Parts * Players(Player) / Points, "0.000"), _
            RoundMid(Parts * Players(Player) / Points, 0), _
            Format((RoundMid(Parts * Players(Player) / Points, 0) - (Parts * Players(Player) / Points)) / (Parts * Players(Player) / Points), "Percent"), _
            Format((RoundMid(Parts * Players(Player) / Points, 0) - 1 - (Parts * Players(Player) / Points)) / (Parts * Players(Player) / Points), "Percent"), _
            Shares(Player)
    Next
    Debug.Print , , , Format(SumSlices, "0.000"), RequestedSlices, , , Parts
    
End Sub

' Demo to run a series of example value sets and list
' the output from RoundingSum.
'
' 2021-03-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RunRoundingSumDemo()

    Dim Values                  As Variant
    Dim Value                   As Variant
    Dim Total                   As Variant
    Dim RequestedTotal          As Variant
    Dim Result                  As Variant
    Dim NumDigitsAfterDecimal   As Long
    Dim ValuesSum               As Variant
    Dim RoundedSum              As Variant
    
    Dim Tests                   As Variant
    Dim Test                    As Integer
    Dim Item                    As Integer
    
    ' Select tests to run.
    Tests = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13)
    
    For Test = LBound(Tests) To UBound(Tests)
        RequestedTotal = 0
        NumDigitsAfterDecimal = 0
        Select Case Tests(Test)
            Case 0
                Values = Array(33, 4, 15, 22, 3, 3, 7, 15, 22, 30, 3, 4, 15, 22, 31, 1, 7, 15, 22, 33, 3, 4, 15, 22, 33, 3, 7, 15, 22, 33, 3, 4, 15, 22, 33, 1, 7, 15, 22, 33)
                RequestedTotal = 4
                NumDigitsAfterDecimal = 0
            Case 1
                Values = Array(-1.66, -1.66, -1.67, 1.7, -1.66)
                RequestedTotal = -11.12
                NumDigitsAfterDecimal = 1
            Case 2
                Values = Array(1.66, 1.66, 1.67, -1.7, 1.66)
                RequestedTotal = -11.12
                NumDigitsAfterDecimal = 1
            Case 3
                Values = Array(1.333333, -1.333333, 1.333333, 2.333333, 1.33)
                RequestedTotal = 0
                NumDigitsAfterDecimal = 1
            Case 4
                Values = Array(1.333333, -1.333333, 1.333333, 2.333333, 1.33)
                RequestedTotal = 5.9
                NumDigitsAfterDecimal = 1
            Case 5
                Values = Array(1.333, 1.333333, 0, 0, 1.33333)
                RequestedTotal = 4.1
            Case 6
                Values = Array(1.333333 * 10 ^ 304, 1.333333 * 10 ^ 304, 0, 1, 1.33 * 10 ^ 304)
                RequestedTotal = 4.1
            Case 7
                Values = Array(433.258, 287.2336, 78.5221, 31198.6551, 4.92236)
                NumDigitsAfterDecimal = -2
            Case 8
                Values = Array(433.258, 287.2336, 78.5221, 31198.6551, 4.92236)
                RequestedTotal = 10000
                NumDigitsAfterDecimal = -2
            Case 9
                Values = Array(433.258, 287.2336, 78.5221, 31198.6551, 4.92236)
                RequestedTotal = 10000
                NumDigitsAfterDecimal = -1
            Case 10
                Values = Array(1432.99999, 2.52, 1.51, 3.55, 0.6)
                RequestedTotal = 0
                NumDigitsAfterDecimal = 0
            Case 11
                Values = Array(1432.99999, -2.52, -1432.99999, 3.55, 2.52, -3.55)
                RequestedTotal = 0
                NumDigitsAfterDecimal = 0
            Case 12
                Values = Array(1.333333 * 10 ^ -304, 1.333333 * 10 ^ -304, 0, 1, 1.33 * 10 ^ -304)
                RequestedTotal = 4.1
                NumDigitsAfterDecimal = 1
            Case 13
                ' Create an array with nine elements to share equally in nine parts.
                Values = Array(9, 9, 9, 9, 9, 9, 9, 9, 9)
                RequestedTotal = 732000
                NumDigitsAfterDecimal = 6
            Case Else
                Values = Null
        End Select
        If Not IsNull(Values) Then
            Debug.Print "Item", "Result  <-", "Input", "Rounded", "Difference", "Weighted Difference"
            Total = RequestedTotal
            ValuesSum = 0
            RoundedSum = 0
            Result = RoundSum(Values, Total, NumDigitsAfterDecimal)
            For Item = LBound(Values) To UBound(Values)
                If Values(Item) = 0 Then
                    Value = 0
                Else
                    Value = Values(Item)
                End If
                Debug.Print _
                    Item, _
                    Result(Item), _
                    Value, _
                    RoundMid(Value, NumDigitsAfterDecimal), _
                    RoundMid(Value, NumDigitsAfterDecimal) - Value, _
                    Value * (RoundMid(Value, NumDigitsAfterDecimal) - Value)
                ValuesSum = ValuesSum + Value
                RoundedSum = RoundedSum + CDbl(RoundMid(Values(Item), NumDigitsAfterDecimal))
            Next
            Debug.Print "Test " & Tests(Test), Total, ValuesSum, RoundedSum
            If RequestedTotal = 0 Then
                RequestedTotal = ValuesSum
            End If
            Debug.Print "Expected:", RoundMid(RequestedTotal, NumDigitsAfterDecimal)
            Debug.Print
        End If
    Next

End Function

' Round the values of one field of a recordset and write the rounded values (matching the total)
' to another field of the recordset.
'
' Example:
'   Rounding the values to two decimals for the sum to match the original sum:
'
'   RoundRecordSum CurrentDb.OpenRecordset("Select Value, RoundedValue From SomeTable"), 0, 2
'
' 2024-03-11. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub RoundRecordSum( _
    ByRef Records As DAO.Recordset, _
    ByVal Total As Double, _
    Optional ByVal NumDigitsAfterDecimal As Long)

    Dim Rows        As Variant
    Dim Values()    As Double
    Dim Value       As Double
    Dim Sum         As Double
    Dim RecordCount As Long
    Dim Index       As Long
    
    If Records.RecordCount = 0 Then
        Exit Sub
    End If
    
    ' Count records.
    Records.MoveLast
    Records.MoveFirst
    RecordCount = Records.RecordCount
    ' Retrieve multi-dimensional array.
    Rows = Records.GetRows(RecordCount)
    
    ' Read values from the first field of the recordset.
    ' Convert to one-dimensional array and calculate the sum.
    ReDim Values(0 To UBound(Rows, 2))
    For Index = LBound(Values) To UBound(Values)
        Value = Rows(0, Index)
        Values(Index) = Value
        Sum = Sum + Value
        Debug.Print Value
    Next
    Debug.Print "Sum:", Sum
    
    ' Adjust sum of output if needed.
    If Total = 0 Then
        Total = Sum
    End If
    
    ' Round the array of values.
    Values = RoundSum(Values, Total, NumDigitsAfterDecimal)
    
    ' Write the values to the second field of the recordset.
    Total = 0
    Records.MoveFirst
    For Index = LBound(Values) To UBound(Values)
        Value = Values(Index)
        If Records(1).Value <> Value Then
            Records.Edit
            Records(1).Value = Value
            Records.Update
        End If
        Total = Total + Value
        Records.MoveNext
    Next
    Debug.Print "Total:", Total
    
End Sub

