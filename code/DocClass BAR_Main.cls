VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
If Selection.Cells.Count > 1 Then Exit Sub
    If BarExists("rClickBar") Then
        CommandBars("rClickBar").ShowPopup
        Cancel = True
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range(rMENU_TYPE)) Is Nothing Then
        Range(rBAR_LOCATION).ClearContents
        Range(rTARGET_CONTROL).ClearContents
    ElseIf Not Intersect(Target, Range(rBAR_LOCATION)) Is Nothing Then
        Range(rTARGET_CONTROL).ClearContents
    End If
End Sub

Public Sub Worksheet_SelectionChange(ByVal Target As Range)
If Selection.Cells.Count > 1 Then Exit Sub
Calculate
OpenValidationComboboxOnClick Target
End Sub

Sub OpenValidationComboboxOnClick(ByVal Target As Range)
    Dim lngValType As Long
    On Error Resume Next
    lngValType = Target.Validation.Type
    On Error GoTo 0
    If lngValType = 3 Then SendKeys "%{DOWN}"
End Sub

