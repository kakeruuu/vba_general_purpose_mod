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

Property Get MergeArrayToRow(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant

    ' ### 結合される側（arr1）が空の場合、結合する方（arr2）をそのまま返す ###
    If IsEmpty(arr1) Then: GoTo errHundle
    
    Dim newRow As Long
    Dim newCol As Long
    
    newRow = UBound(arr1, 1) + UBound(arr2, 1) + 1
    newCol = Application.WorksheetFunction.Max(UBound(arr1, 2), UBound(arr2, 2))
    
    Dim newArr As Variant
    ReDim newArr(0 To newRow, 0 To newCol)
    
    
    Dim i As Long
    Dim j As Long
    
    For i = 0 To newRow
        If i <= UBound(arr1, 1) Then
            For j = 0 To newCol
                If j <= UBound(arr1, 2) Then
                    newArr(i, j) = arr1(i, j)
                Else
                    newArr(i, j) = Empty
                End If
            Next j
        Else
            For j = 0 To newCol
                If j <= UBound(arr2, 2) Then
                    newArr(i, j) = arr2(i - UBound(arr1, 1) - 1, j)
                Else
                    newArr(i, j) = Empty
                End If
            Next j
        End If
    Next i
    
    MergeArrayToRow = newArr
    
    Exit Property
    
errHundle:
    MergeArrayToRow = arr2
End Property

Property Get MergeArrayToCol()
    
    Dim newRow As Long
    Dim newCol As Long
    
    newRow = Application.WorksheetFunction.Max(UBound(arr1, 1), UBound(arr2, 1))
    newCol = UBound(arr1, 2) + UBound(arr2, 2)
    
    Dim newArr As Variant
    ReDim newArr(0 To newRow, 0 To newCol)
    
    
    Dim i As Long
    Dim j As Long
    For j = 0 To newCol
        If j <= UBound(arr1, 2) Then
            For i = 0 To newRow
                If i <= UBound(arr1, 1) Then
                    newArr(i, j) = arr1(i, j)
                Else
                    newArr(i, j) = Empty
                End If
            Next j
        Else
            For i = 0 To newRow
                If j <= UBound(arr2, 1) Then
                    newArr(i, j) = arr2(i - UBound(arr1, 1) - 1, j)
                Else
                    newArr(i, j) = Empty
                End If
            Next j
        End If
    Next i
    
    MergeArrayToRow = newArr
    
End Property
Property Get RemoveColumnArray2D(ByRef arr As Variant, ByRef delCol As Long)
    
    Dim rMin As Long: rMin = LBound(arr, 1)
    Dim rMax As Long: rMax = LBound(arr, 1)
    Dim cMin As Long: cMin = LBound(arr, 1)
    Dim cMax As Long: cMax = LBound(arr, 1)
    
    Dim temp As Variant
    ReDim temp(rMin To Max, cMin To cMax - 1)
    
    Dim R As Long               '二次元配列のRow要素
    Dim C As Long               '二次元配列のColumn要素
    Dim i As Long: i = rMin     '二次元配列のColumn要素(元配列arr)
    
    
    For R = rMin To rMax
        For C = cMin To cMax - 1
            If C = delCol Then
                i = i + 1
            End If
            temp(R, C) = arr(R, i)
            i = i + 1
        Next C
            i = rMin
    Next R
    
    RemoveColumnArray2D = temp
    
End Property
