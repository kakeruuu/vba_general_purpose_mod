VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetWorkBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private curPath As String
Private oxl As Excel.Application
Private owb As Excel.Workbooks
    
    
Private Sub Class_Initialize()

    curPath = ThisWorkbook.Path
        
    Set oxl = CreateObject("Excel.Application")
    
    Set owb = oxl.Application.Workbooks
    
End Sub

Public Function AbsolutelyPath(tergetFileName) As String

    AbsolutelyPath = curPath & "\" & tergetFileName
    
End Function

Public Function OpenBook(tergetFileName) As Workbook

    Set OpenBook = Workbooks.Open(Me.AbsolutelyPath(tergetFileName))
    
End Function

Public Function HiddenOpenBook(tergetFileName) As Object

    Set wbReadBook = owb.Open(Me.AbsolutelyPath(tergetFileName))
    
End Function
