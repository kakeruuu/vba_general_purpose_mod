VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetWorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private thisWb As Workbook
Public thisSh As Worksheet
Private thisShName As String

Public Function Init(ByRef tergetWb As Workbook, ByVal shName As String) As SetWorkSheet

    Set Init = Me
    Set thisWb = tergetWb
    Set thisSh = tergetWb.Sheets(shName)
    thisShName = shName
    
End Function


' �^����ꂽ�͈͂̒l��z��ɂ��ĕԂ�
Public Function RangeToArray(ByRef tergetRange As Range) As Variant

    Dim rangeThisSh As Range
    Set rangeThisSh = thisSh.Range(tergetRange.Address)
    
    Dim rangeAry() As Variant
    ReDim rangeAry(rangeThisSh.Rows.Count - 1, rangeThisSh.Columns.Count - 1)
    
    Dim row As Range
    Dim column As Range
    Dim i As Long: i = 0
    Dim j As Long: j = 0
    
    For Each row In rngThisSh.Rows
        For Each column In row.Columns
            rangeAry(i, j) = column.Value
            j = j + 1
        Next column
        j = 0
        i = i + 1
    Next row
    
    RangeToArray = rangeAry
    
End Function