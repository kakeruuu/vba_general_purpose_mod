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
