Attribute VB_Name = "main"
Option Explicit

Sub loopWrapperForWorkSheets(func)
    ''' Activate����Ă���Workbook���̂��ׂẴV�[�g�Ɉ����̊֐������s���� '''
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim sheets As Worksheets
    Set sheets = wb.Worksheets
    
    Dim sh As Worksheet
    For Each sh In sheets
        sh.Activate
        func
    Next sh
    
End Sub
