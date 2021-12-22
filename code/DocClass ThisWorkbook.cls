VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_AddinInstall()
CreateAllBars
End Sub

Private Sub Workbook_AddinUninstall()
DeleteAllBars
End Sub

Public Sub Workbook_BeforeClose(Cancel As Boolean)
DeleteAllBars
End Sub

Public Sub Workbook_Open()
CreateAllBars
'runListProcs
End Sub

