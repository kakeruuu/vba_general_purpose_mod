VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassArray"
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
    
    newRow = UBound(arr1, 1) + UBound(arr2, 1)
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
    
    Exit Property
    
errHundle:
    MergeArrayToRow = arr2
End Property

Property Get MergeArrayToCol(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant
    
    If IsEmpty(arr1) Then: GoTo errHundle
    
    Dim newRow As Long
    Dim newCol As Long
    
    newRow = Application.WorksheetFunction.Max(UBound(arr1, 1), UBound(arr2, 1))
    newCol = UBound(arr1, 2) + UBound(arr2, 2)
    
    Dim newArr As Variant
    ReDim newArr(1 To newRow, 1 To newCol)
    
    
    Dim i As Long
    Dim j As Long
    For j = 1 To newCol
        If j <= UBound(arr1, 2) Then
            For i = 1 To newRow
                If i <= UBound(arr1, 1) Then
                    newArr(i, j) = arr1(i, j)
                Else
                    newArr(i, j) = Empty
                End If
            Next i
        Else
            For i = 1 To newRow
                If i <= UBound(arr2, 1) Then
                    newArr(i, j) = arr2(i, j - UBound(arr1, 2))
                Else
                    newArr(i, j) = Empty
                End If
            Next i
        End If
    Next j
    
    MergeArrayToCol = newArr
    
    Exit Property
    
errHundle:
    MergeArrayToCol = arr2
    
End Property
' arr[0,1,2,3,4] → arr[[0],[1],[2],[3],[4]]
Property Get ArrayTo2dRowArray(ByRef arr As Variant) As Variant
    
    Dim buf() As Variant: ReDim buf(1 To UBound(arr), 1)
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        buf(i, LBound(arr)) = arr(i)
    Next i
    
    ArrayTo2dRowArray = buf
    
End Property

' arr[0,1,2,3,4] → arr[[0,1,2,3,4]]
Property Get ArrayTo2dColArray(ByRef arr As Variant) As Variant
    
    Dim buf() As Variant: ReDim buf(1 To LBound(arr, 1), 1 To UBound(arr))
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        buf(LBound(arr), i) = arr(i)
    Next i
    
    ArrayTo2dColArray = buf
    
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

' [[0,1,2],[3,4,5]] → {0: [0,3], 1:[1,4], 2:[2,5]}
' keyは列番号のイメージ
Public Function Array2DToDict(ByRef arr As Variant) As Object
    
    Dim dic As New Dictionary
    Dim i As Long
    Dim j As Long
    
    Dim R As Long: R = 1
    Dim newArr As Variant
    For j = 1 To UBound(arr, 2)
        ReDim newArr(1 To UBound(arr, 1))
        For i = 1 To UBound(arr, 1)
            newArr(i) = arr(i, j)
        Next i
        dic.Add R, newArr
        R = R + 1
    Next j
    
    Set Array2DToDict = dic
    
End Function

' memo:今後Dictionaryクラスを作ったらそっちに移行する予定
' {0: [0,3], 1:[1,4], 2:[2,5]} → [[0,1,2],[3,4,5]]
Property Get DictToArray2D(ByRef dic As Dictionary) As Variant
    Dim keys() As Variant
    keys = dic.keys
    
    Dim rowsAryInDict As Long: rowsAryInDict = UBound(dic.Item(keys(0)))
    Dim newArr() As Variant: ReDim newArr(1 To rowsAryInDict, 1 To UBound(keys, 1) + 1)
    
    Dim i As Long
    Dim j As Long
    Dim aryInDict As Variant
    For i = 0 To UBound(keys, 1)
        aryInDict = dic.Item(keys(i))
        For j = 1 To rowsAryInDict
            newArr(j, i + 1) = aryInDict(j)
        Next j
    Next i
    
    DictToArray2D = newArr
End Property

' 配列の特定の列に対してある値が含まれている確認し、入っていたらその行を削除する関数
Property Get ExcludeRowForIncludeWordInArray(ByRef ary As Variant, ByVal tergetCol As Long, ByVal tergetWord As Variant) As Variant
    
    Dim newArr() As Variant: ReDim newArr(1 To LBound(ary, 1), 1 To UBound(ary, 2))
    Dim sumArr As Variant
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To UBound(ary, 1)
        ' 配列行内のtergetColにtergetWordが含まれていなかったら
        If InStr(ary(i, tergetCol), tergetWord) = 0 Then
            For j = 1 To UBound(ary, 2)
                newArr(1, j) = ary(i, j)
            Next j
            ' 空白行を詰める処理を行いたくないため1行ずつマージする処理にしている
            sumArr = Me.MergeArrayToRow(sumArr, newArr)
        End If
    Next i
    
    ExcludeRowForIncludeWordInArray = sumArr
    
End Property
