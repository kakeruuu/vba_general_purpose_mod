VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassWorkBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private curPath As String
Private oxl As Excel.Application
Private owb As Excel.Workbooks
    
' ���̃u�b�N����ɑΏۂ̃u�b�N��I�ԁB
' ���̃u�b�N����̊K�w�̃u�b�N�͑���ł��Ȃ����Afso���g���Ӗ����������߈�UThisWorkBook�𗘗p���Ă���
Private Sub Class_Initialize()

    curPath = ThisWorkbook.Path
    
    Set oxl = CreateObject("Excel.Application")
    
    Set owb = oxl.Application.Workbooks
    
End Sub

Property Get AbsolutelyPath(tergetFileName) As String
    AbsolutelyPath = curPath & "\" & tergetFileName
End Property

Property Get OpenBook(tergetFileName) As Workbook
    Set OpenBook = Workbooks.Open(Me.AbsolutelyPath(tergetFileName))
End Property

Property Get HiddenOpenBook(tergetFileName) As Object
    Set HiddenOpenBook = owb.Open(Me.AbsolutelyPath(tergetFileName))
End Property