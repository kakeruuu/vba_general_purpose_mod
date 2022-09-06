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

Private tergetWb
Private shName

Public Function Init(ByRef tergetWb As Workbook, ByVal shName As String) As SetWorkSheet

    Set Init = Me
    tergetWb = tergetWb
    shName = shName
    
End Function


' —^‚¦‚ç‚ê‚½”ÍˆÍ‚Ì’l‚ð”z—ñ‚É‚µ‚Ä•Ô‚·
Public Function RangeToArray(ByRef tergetRng As Range)
    
End Function

