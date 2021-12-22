Attribute VB_Name = "Z_Cell_Groupbox_Menu"
Sub MoveShapes()
    'add to worksheet selection change

    If Target.Cells.Count > 1 Then Exit Sub
    With ActiveWindow.VisibleRange
        ActiveSheet.Shapes("PictureX").Top = .Top + 5
        ActiveSheet.Shapes("PictureX").Left = .Left + .Width - ActiveSheet.Shapes("PictureX").Width - 45
        ActiveSheet.Shapes("PictureX").Top = .Top + 35
        ActiveSheet.Shapes("PictureX").Left = .Left + .Width - ActiveSheet.Shapes("PictureX").Width - 45
    End With
End Sub

Sub FormControl()
    On Error Resume Next
    ActiveSheet.Shapes("grpCustomControls").Delete
    Err.Clear: On Error GoTo -1: On Error GoTo 0
    With ActiveSheet.GroupBoxes.Add(Range("D4").Left, Range("D4").Top, Range("D4:G4").Width, Range("D4:D6").Height)
        .Characters.Text = "Custom Tools"
        .Name = "grpBox"
    End With
 
    With ActiveSheet.Buttons.Add(Range("D4").Left, Range("D5").Top, Range("D4").Width, Range("D5:D6").Height)
        .Name = "btnOne"
        .Characters.Text = "Add New"
        .OnAction = "test"
    End With
    With ActiveSheet.Buttons.Add(Range("E4").Left, Range("E5").Top, Range("E4").Width, Range("E5:E6").Height)
        .Name = "btnTwo"
        .Characters.Text = "Goto 1st Rec"
        .OnAction = "test"
    End With
    With ActiveSheet.Buttons.Add(Range("F4").Left, Range("F5").Top, Range("F4").Width, Range("F5:F6").Height)
        .Name = "btnThree"
        .Characters.Text = "Remove Filters"
        .OnAction = "test"
    End With
    With ActiveSheet.Buttons.Add(Range("G4").Left, Range("G5").Top, Range("G4").Width, Range("G5:G6").Height)
        .Name = "btnFour"
        .Characters.Text = "Copy W/S"
        .OnAction = "test"
    End With
    ActiveSheet.Shapes.Range(Array("grpBox", "btnOne", "btnTwo", "btnThree", "btnFour")).Select
    With Selection.ShapeRange.Group
        .Name = "grpCustomControls"
    End With
 
End Sub

Sub DeleteFormControl()
    ActiveSheet.Shapes("grpCustomControls").Delete
End Sub

