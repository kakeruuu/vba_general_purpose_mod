VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function MergeArrayToRow(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant

    Dim newRow As Long
    Dim newCol As Long
    
    newRow = UBound(arr1, 1) + UBound(arr2, 2)
    newCol = Application.WorksheetFunction.Max(UBound(arr1, 2), UBound(arr2, 2))
    
    Dim newArr As Variant
    ReDim newArr(1 To newRow, 1 To newCol)
    
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To newRow
        If i <= UBound(arr1, 1) Then
            For j = 1 To newCol
                If j <= UBound(arr1, 2) Then
                    newArr(i, j) = arr1(i, j)
                Else
                    newArr(i, j) = Empty
                End If
            Next j
        Else
            For j = 1 To newCol
                If j <= UBound(arr2, 2) Then
                    newArr(i, j) = arr2(i - UBound(arr1, 1), j)
                Else
                    newArr(i, j) = Empty
                End If
            Next j
        End If
    Next i
    
    MergeArrayToRow = newArr
    
End Function
