Attribute VB_Name = "main"
Option Explicit

Sub loopWrapperForWorkSheets(func)
    ''' ActivateされているWorkbook内のすべてのシートに引数の関数を実行する '''
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

Sub loopWrapperForWorkbooks(func, dir_name)
    
    Dim fso As New FileSystemObject
    Dim tergetDir As Folder
    
    Set tergetDir = fso.GetFolder(dir_name)
    Dim tergetFile As file
    
    For Each tergetFile In tergetDir.Files
        currentFile = Workbooks.Open(tergetFile.Name)
        func
    Next tergetFile
    
End Sub
    
