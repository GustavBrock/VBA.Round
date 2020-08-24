Attribute VB_Name = "ArraySorting"
' ArraySorting v1.1.0
' (c) 2018-03-26. Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Round
'
' Set of functions for sorting arrays.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Explicit

' Quickly sort a Variant array.
'
' The array does not have to be zero- or one-based.
'
' 2018-03-16. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub QuickSort(ByRef Values As Variant)

    Dim Lows()      As Variant
    Dim Mids()      As Variant
    Dim Tops()      As Variant
    Dim Pivot       As Variant
    Dim Lower       As Long
    Dim Upper       As Long
    Dim UpperLows   As Long
    Dim UpperMids   As Long
    Dim UpperTops   As Long
    
    Dim Value       As Variant
    Dim Item        As Long
    Dim Index       As Long
 
    ' Find count of elements to sort.
    Lower = LBound(Values)
    Upper = UBound(Values)
    If Lower = Upper Then
        ' One element only.
        ' Nothing to do.
        Exit Sub
    End If
    
    
    ' Choose pivot in the middle of the array.
    Pivot = Values(Int((Upper - Lower) / 2) + Lower)
    ' Construct arrays.
    For Each Value In Values
        If Value < Pivot Then
            ReDim Preserve Lows(UpperLows)
            Lows(UpperLows) = Value
            UpperLows = UpperLows + 1
        ElseIf Value > Pivot Then
            ReDim Preserve Tops(UpperTops)
            Tops(UpperTops) = Value
            UpperTops = UpperTops + 1
        Else
            ReDim Preserve Mids(UpperMids)
            Mids(UpperMids) = Value
            UpperMids = UpperMids + 1
        End If
    Next
    
    ' Sort the two split arrays, Lows and Tops.
    If UpperLows > 0 Then
        QuickSort Lows()
    End If
    If UpperTops > 0 Then
        QuickSort Tops()
    End If
    
    ' Concatenate the three arrays and return Values.
    Item = 0
    For Index = 0 To UpperLows - 1
        Values(Lower + Item) = Lows(Index)
        Item = Item + 1
    Next
    For Index = 0 To UpperMids - 1
        Values(Lower + Item) = Mids(Index)
        Item = Item + 1
    Next
    For Index = 0 To UpperTops - 1
        Values(Lower + Item) = Tops(Index)
        Item = Item + 1
    Next

End Sub

' Demonstrates the usage of function QuickSort.
'
' 2018-03-16. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub QuickSortTest()

    Dim Samples(1 To 26)    As Variant
    Dim Item                As Long
 
    ' Populate Samples with numbers in descending order.
    For Item = 1 To 26: Samples(Item) = 26 - Item: Next
    For Item = 1 To 26: Debug.Print Samples(Item);: Next
    Debug.Print
    
    ' Sort ascending.
    QuickSort Samples()
    For Item = 1 To 26: Debug.Print Samples(Item);: Next
    Debug.Print
    
    ' Populate Samples with strings in descending order.
    For Item = 1 To 26: Samples(Item) = Chr(Asc("z") + 1 - Item) & "-item": Next
    For Item = 1 To 26: Debug.Print Samples(Item); " ";: Next
    Debug.Print
    
    ' Sort ascending.
    QuickSort Samples()
    For Item = 1 To 26: Debug.Print Samples(Item); " ";: Next
    Debug.Print

End Sub

' Quickly fill an array with the index of the sorting order of another array.
'
' The arrays do not have to be zero- or one-based.
' For typical usage, see function QuickSortIndexTest.
'
' 2018-03-16. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub QuickSortIndex( _
    ByRef Pointers() As Long, _
    ByVal Values As Variant)

    Dim Lows()      As Long
    Dim Mids()      As Long
    Dim Tops()      As Long
    
    Dim UpperLows   As Long
    Dim UpperMids   As Long
    Dim UpperTops   As Long
    
    Dim Lower       As Long
    Dim Upper       As Long
    
    Dim Pivot       As Long
    Dim Item        As Long
    Dim Index       As Long
    Dim Pointer     As Variant
 
    ' Find count of elements to sort.
    Lower = LBound(Pointers)
    Upper = UBound(Pointers)
    If Lower = Upper Then
        ' One element only.
        ' Nothing to do.
        Exit Sub
    End If
    
    If Pointers(Lower) = Pointers(Upper) Then
        ' Fill array with default sorting.
        For Item = Lower To Upper
            Pointers(Item) = Item
        Next
    End If
    
    ' Choose pivot as the the middle of the array.
    Pivot = Pointers(Int((Upper - Lower) / 2) + Lower)
    ' Construct arrays.
    For Each Pointer In Pointers
        If Values(Pointer) < Values(Pivot) Then
            ReDim Preserve Lows(UpperLows)
            Lows(UpperLows) = Pointer
            UpperLows = UpperLows + 1
        ElseIf Values(Pointer) > Values(Pivot) Then
            ReDim Preserve Tops(UpperTops)
            Tops(UpperTops) = Pointer
            UpperTops = UpperTops + 1
        Else
            ReDim Preserve Mids(UpperMids)
            Mids(UpperMids) = Pointer
            UpperMids = UpperMids + 1
        End If
    Next
    
    ' Sort the two split arrays, Lows and Tops.
    If UpperLows > 0 Then
        QuickSortIndex Lows(), Values
    End If
    If UpperTops > 0 Then
        QuickSortIndex Tops(), Values
    End If
    
    ' Concatenate the three arrays and return array Pointers.
    Item = 0
    For Index = 0 To UpperLows - 1
        Pointers(Lower + Item) = Lows(Index)
        Item = Item + 1
    Next
    For Index = 0 To UpperMids - 1
        Pointers(Lower + Item) = Mids(Index)
        Item = Item + 1
    Next
    For Index = 0 To UpperTops - 1
        Pointers(Lower + Item) = Tops(Index)
        Item = Item + 1
    Next

End Sub

' Demonstrates the usage of function QuickSortIndex.
'
' 2018-03-16. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub QuickSortIndexTest()

    Dim Pointers()      As Long
    
    Dim Values(4 To 11) As Variant
    Dim Index           As Variant
    
    ' Fill values.
    For Index = LBound(Values) To UBound(Values)
        Values(Index) = Array("Don", "Kit", "Zoe", "Il", "Wue", "Onu", "Bo", "Ann")(Index - LBound(Values))
    Next
    ReDim Pointers(LBound(Values) To UBound(Values))

    ' Return Pointers with the sorting order of Values.
    QuickSortIndex Pointers(), Values
    
    Debug.Print
    For Index = LBound(Values) To UBound(Values)
        Debug.Print Index, Values(Index), Pointers(Index), Values(Pointers(Index))
    Next
    
End Sub

' Reverse the order of items of an array.
'
' 2018-03-16. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub ReverseSort(ByRef Values As Variant)

    Dim Reversed    As Variant
    Dim Item        As Long
    
    ' Create a copy of the array.
    Reversed = Values
    
    ' Fill the array with the items in reverse order.
    For Item = LBound(Values) To UBound(Values)
        Reversed(UBound(Values) - Item + LBound(Values)) = Values(Item)
    Next
    
    ' Return the reversed array.
    Values = Reversed
    
End Sub

