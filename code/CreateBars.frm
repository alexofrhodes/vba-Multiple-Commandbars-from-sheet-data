VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateBars 
   Caption         =   "RaiseTheBar ~ by Anastasiou Alex"
   ClientHeight    =   5664
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5712
   OleObjectBlob   =   "CreateBars.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public validTag As Boolean

Private Sub BarTag_Change()
    validTag = False
    BarTag.ForeColor = vbRed
    
    If BarList.List(BarList.ListIndex) = "VBEMenu" Then
        BarLocation.List(BarLocation.ListCount - 1) = "Floating -" & BarTag.Text & "-"
    End If
    vbeBarControls.List(0) = "-" & BarTag.Text & "- "


    If BarTag.Text = "" Or IsNumeric(BarTag.Text) Then Exit Sub
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("combarTAGS").Columns(1).Find(BarTag.Text, , lookat:=xlWhole)
    If rng Is Nothing Then
        Set rng = ThisWorkbook.Sheets("combarLIST").Columns(1).Find(BarTag.Text, , lookat:=xlWhole)
        If rng Is Nothing Then validTag = True
        BarTag.ForeColor = vbBlue
    End If
End Sub

Private Sub UserForm_Initialize()
    BarList.AddItem "WorksheetMenu"
    BarList.AddItem "VBEMenu"
    BarList.AddItem "RightClickMenu"
    BarList.ListIndex = 0
    
    vbeBarControls.AddItem IIf(BarTag.Text = "", "-TAG-", "-" & BarTag.Text & "-")
    vbeBarControls.AddItem "File"
    vbeBarControls.AddItem "Edit"
    vbeBarControls.AddItem "View"
    vbeBarControls.AddItem "Insert"
    vbeBarControls.AddItem "Format"
    vbeBarControls.AddItem "Debug"
    vbeBarControls.AddItem "Run"
    vbeBarControls.AddItem "Tools"
    vbeBarControls.AddItem "Add-Ins"
    vbeBarControls.AddItem "Window"
    vbeBarControls.AddItem "Help"
End Sub

Private Sub BarList_Click()
            CommandButton1.Enabled = True

    Select Case BarList.List(BarList.ListIndex)
    Case Is = "WorksheetMenu"
        BarLocation.Clear
        BarLocation.AddItem "Worksheet Menu Bar"
        BarLocation.AddItem "Cell"
        BarLocation.AddItem "Column"
        BarLocation.AddItem "Row"
    Case Is = "VBEMenu"
        BarLocation.Clear
        BarLocation.AddItem "Menu Bar"
        BarLocation.AddItem "Code Window"
        BarLocation.AddItem "Project Window"
        BarLocation.AddItem "Edit"
        BarLocation.AddItem "Debug"
        BarLocation.AddItem "Userform"
        BarLocation.AddItem IIf(BarTag.Text = "", "Floating -TAG-", "-" & BarTag.Text & "-")
    Case Else
        BarLocation.Clear
        If WorksheetFunction.CountIf(Sheets("RaiseTheBar").Columns(1), "1") > 1 Then
            MsgBox "Cant't create more than 1 -level 1 menu- for right click popup bars. Create separate bars if more than one are needed."
            CommandButton1.Enabled = False
            vbeBarControls.Visible = False

        Exit Sub
        End If
    End Select
    
    Select Case BarList.ListIndex
    Case Is = 0, 1
        BarLocation.Visible = True
        BarLocation.ListIndex = 0
    Case Else
        BarLocation.Visible = False
    End Select
    
    Select Case BarList.ListIndex
    Case Is = 1
        vbeBarControls.ListIndex = 0
    End Select
    
    Select Case BarList.ListIndex
    Case Is = 0, 2
        vbeBarControls.Visible = False
    End Select
End Sub

Private Sub BarLocation_Click()
    If BarList.ListIndex = 1 And BarLocation.ListIndex = 0 Then
        vbeBarControls.Visible = True
    Else
        vbeBarControls.Visible = False
    End If
End Sub

Private Sub CommandButton1_Click()
    If validTag = False Then
        MsgBox "Tag must be Unique"
        Exit Sub
    End If
    If BarList.ListIndex = -1 Or _
       (BarList.ListIndex < BarList.ListCount - 1 And BarLocation.ListIndex = -1) Then
        MsgBox "Fill required fields"
        Exit Sub
    End If
    CommandBarBuilder
End Sub


