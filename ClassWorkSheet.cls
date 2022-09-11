VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassWorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private thisWb As Workbook
Public thisSh As Worksheet
Private thisShName As String

Public Function Init(ByRef tergetWb As Workbook, ByVal shName As String) As ClassWorkSheet

    Set Init = Me
    Set thisWb = tergetWb
    ' memo:thisShでSheetオブジェクトを返すか、OpenSheetsのようなプロシージャを作るべきか
    ' 後者の方が変数にシートオブジェクトを代入するというイメージが沸きやすいが
    ' 特定のシートを操作するというより、シートオブジェクト全般を操作するためのクラスとしての意味合いが強くなる？
    Set thisSh = tergetWb.Sheets(shName)
    thisShName = shName
    
End Function

' シート内のデータが存在する範囲の最終行を返す
Property Get UsedEndRow() As Long
    UsedEndRow = thisSh.UsedRange.Rows.Count
End Property

Property Get UsedEndCol() As Long
    UsedEndCol = thisSh.UsedRange.Columns.Count
End Property

' 与えられた範囲の値を配列にして返す
Public Function RangeToArray(ByRef tergetRange As range) As Variant

    Dim rangeThisSh As range
    Set rangeThisSh = thisSh.range(tergetRange.Address)
    
    Dim rangeAry() As Variant
    ReDim rangeAry(rangeThisSh.Rows.Count - 1, rangeThisSh.Columns.Count - 1)
    
    Dim row As range
    Dim column As range
    Dim i As Long: i = 0
    Dim j As Long: j = 0
    
    For Each row In rangeThisSh.Rows
        For Each column In row.Columns
            rangeAry(i, j) = column.Value
            j = j + 1
        Next column
        j = 0
        i = i + 1
    Next row
    
    RangeToArray = rangeAry
    
End Function
