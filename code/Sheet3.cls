VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
If Selection.Cells.Count > 1 Then Exit Sub
    If BarExists("rClickBar") Then
        CommandBars("rClickBar").ShowPopup
        Cancel = True
    End If
End Sub




Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Selection.Cells.Count > 1 Then Exit Sub

ActionValidation Target

OpenValidationComboboxOnClick Target
End Sub

Sub OpenValidationComboboxOnClick(ByVal Target As Range)
If Intersect(Target, Columns("C:D")) Is Nothing Then Exit Sub
    Dim lngValType As Long
    On Error Resume Next
    lngValType = Target.Validation.Type
    On Error GoTo 0
    If lngValType = 3 Then SendKeys "%{DOWN}"
End Sub

Sub ActionValidation(ByVal Target As Range)
If Intersect(Target, Columns("C")) Is Nothing Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Procs")
    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion
    With Target.Validation
        .Delete 'delete previous validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="='" & ws.Name & "'!" & rng.Address
    End With
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
LevelValidation Target
End Sub
Sub LevelValidation(ByVal Target As Range)
If Selection.Cells.Count > 1 Then Exit Sub
If Intersect(Target, Columns("A")) Is Nothing Then Exit Sub
If Target = "" Or Target = 1 Then Exit Sub
Select Case Target
Case Is > 3, Is < 1
Target = ""
MsgBox "Levels allowed 1-3"
Exit Sub
Case Else
Select Case Target - Target.Offset(-1, 0)
Case Is = 0, 1
'nothing
Case Else
Target = ""
MsgBox "Levels allowed = 1, = previous, = previous + 1"
End Select
End Select

End Sub
