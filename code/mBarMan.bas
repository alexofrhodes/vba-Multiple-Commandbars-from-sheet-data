Attribute VB_Name = "mBarMan"
Private Const rBUILD_ON_OPEN = "I2"
Private Const rC_TAG = "I4"
Private Const rMENU_TYPE = "I5"
Private Const rBAR_LOCATION = "I6"
Private Const rTARGET_CONTROL = "I7"
Private C_TAG As String
Private MenuEvent As CVBECommandHandler
Private EventHandlers As New Collection
Private CmdBarItem As CommandBarControl
Private TargetCommandbar As CommandBar
Private TargetControl As CommandBarControl
Private MainMenu As CommandBarControl
Private MenuItem As CommandBarControl
Private Ctrl As Office.CommandBarControl
Private MenuLevel, NextLevel, Caption, Divider, FaceId
Private Action As String
Private MenuSheet As Worksheet
Private row As Integer
Private MenuType As Long
Private Const WorksheetMenu = 1
Private Const VbeMenu = 2
Private Const RightClickMenu = 3
Private BarLocation As String

Public Sub CreateAllBars()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(Left(ws.Name, 4)) = "BAR_" Then
            If ws.Range(rBUILD_ON_OPEN) = True Then CommandBarBuilder ws
        End If
    Next
End Sub

Public Sub DeleteAllBars()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(Left(ws.Name, 4)) = "BAR_" Then DeleteControlsAndHandlers ws
    Next
End Sub

Public Sub RestoreBars()
    Application.OnTime Now, "CreateAllBars"
End Sub

Public Sub ListBars()
    ListWorksheetBars
    ListVBEBars
End Sub

Public Sub NewBar()
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets("BAR_Main")
    Dim wsCopy As Worksheet
    wsMain.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    Set wsCopy = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    wsCopy.Name = "BAR_" & lastBar + 1
    wsCopy.Range("A1").CurrentRegion.Offset(1).ClearContents
    wsCopy.Range("I4:I7").ClearContents
    wsCopy.Range("I2") = False
End Sub

Private Function lastBar() As Long
    Dim counter As Long
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(Left(ws.Name, 4)) = "BAR_" Then counter = counter + 1
    Next
    lastBar = counter
End Function

Private Function SetCMDbar(ws As Worksheet) As Boolean
    C_TAG = ws.Range(rC_TAG)
    Select Case LCase(ws.Range(rMENU_TYPE))
        Case Is = LCase("WorksheetMenu")
            MenuType = WorksheetMenu
        Case Is = LCase("vbeMenu")
            MenuType = VbeMenu
        Case Is = LCase("RightClickMenu")
            MenuType = RightClickMenu
        Case Else
    End Select
    If ws.Range(rBAR_LOCATION) <> "" Then
        BarLocation = ws.Range(rBAR_LOCATION)
    Else
        BarLocation = 0
    End If
    If MenuType = VbeMenu Then
        Select Case BarLocation
            Case Is = "Menu Bar", "Code Window", "Project Window", "Edit", "Debug", "Userform"
                Set TargetCommandbar = Application.VBE.CommandBars(BarLocation)
                SetCMDbar = True
            Case Else
                Set TargetCommandbar = Application.VBE.CommandBars.Add(C_TAG, position:=msoBarTop, Temporary:=True)
                TargetCommandbar.Visible = True
        End Select
    ElseIf MenuType = WorksheetMenu Then
        Select Case ws.Range(rBAR_LOCATION)
            Case Is = "Worksheet Menu Bar", "Cell", "Column", "Row"
                Set TargetCommandbar = Application.CommandBars(BarLocation)
                SetCMDbar = True
            Case Else
        End Select
    Else
    End If
End Function

Public Function BarExists(findBarName As String) As Boolean
    Dim bar As CommandBar
    For Each bar In Application.CommandBars
        If bar.Name = findBarName Then
            BarExists = True
            Exit Function
        End If
    Next bar
    For Each bar In Application.VBE.CommandBars
        If bar.Name = findBarName Then
            BarExists = True
            Exit Function
        End If
    Next bar
End Function

Public Sub BuildBarFromShape()
    CommandBarBuilder ActiveSheet
End Sub

Public Sub DeleteBarFromShape()
    DeleteControlsAndHandlers ActiveSheet
End Sub

Private Sub CommandBarBuilder(ws As Worksheet)
    DeleteControlsAndHandlers ws
    SetCMDbar ws
    Set MenuSheet = ws
    row = 2
    If MenuType = VbeMenu Then
        If BarLocation = "Menu Bar" Then
            Select Case LCase(ws.Range(rTARGET_CONTROL))
                Case LCase(ws.Range(rC_TAG))
                    Set TargetControl = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
                Case Else
                    Dim vbControl As String
                    vbControl = ws.Range(rTARGET_CONTROL)
                    Set TargetControl = TargetCommandbar.Controls(vbControl).Controls.Add(Type:=msoControlPopup, Temporary:=True)
            End Select
        Else
            If LCase(ws.Range(rTARGET_CONTROL)) <> LCase(ws.Range(rC_TAG)) _
        And ws.Range(rTARGET_CONTROL) <> "" Then _
                                      Set TargetControl = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
        End If
        If Not TargetControl Is Nothing Then
            TargetControl.Caption = C_TAG
            TargetControl.Tag = C_TAG
        End If
    End If
    Do Until IsEmpty(MenuSheet.Cells(row, 1))
        With MenuSheet
            SetVariables
        End With
        Select Case MenuLevel
            Case 1
                If NextLevel > MenuLevel Then
                    CreateMainMenu
                Else
                    DirectButton
                End If
            Case 2
                If NextLevel > MenuLevel Then
                    CreatePopup
                Else
                    DirectButton
                End If
            Case 3
                CreateButton
        End Select
        row = row + 1
        ReSetVariables
    Loop
    markControlType ws
    Debug.Print "Bar created"
End Sub

Private Sub markControlType(ws As Worksheet)
    ws.Columns("F").ClearContents
    Dim idx As Long: idx = 0
    Dim Description() As Variant
    Dim cell As Range
    Set cell = ws.Cells(2, 1)
    Do Until IsEmpty(cell)
        idx = idx + 1
        ReDim Preserve Description(1 To idx)
        Description(idx) = IIf(cell.Offset(1) > cell, "PopUp", "Button")
        Set cell = cell.Offset(1)
    Loop
    ws.Range("F2").Resize(UBound(Description)) = WorksheetFunction.Transpose(Description)
End Sub

Private Sub SetVariables()
    With MenuSheet
        MenuLevel = .Cells(row, 1)
        Caption = .Cells(row, 2)
        Action = .Cells(row, 3)
        Divider = .Cells(row, 4)
        FaceId = .Cells(row, 5)
        NextLevel = .Cells(row + 1, 1)
    End With
End Sub

Private Sub ReSetVariables()
    MenuLevel = ""
    Caption = ""
    Action = ""
    Divider = ""
    FaceId = ""
    NextLevel = ""
End Sub

Private Sub CreateMainMenu()
    If MenuType = VbeMenu Then
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup)
    ElseIf MenuType = WorksheetMenu Then
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    ElseIf MenuType = RightClickMenu Then
        Set TargetCommandbar = CommandBars.Add(C_TAG, msoBarPopup, , True)
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup)
        'Exit Sub
    End If
    With MainMenu
        .Caption = Caption
        .BeginGroup = Divider
        If FaceId <> "" And Action <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
End Sub

Private Sub CreatePopup()
    If MenuType = RightClickMenu Then
        Set MenuItem = TargetCommandbar.Controls.Add(Type:=msoControlPopup)
    Else
        Set MenuItem = MainMenu.Controls.Add(Type:=msoControlPopup)
    End If
    With MenuItem
        .Caption = Caption
        .BeginGroup = Divider
        If FaceId <> "" And Action <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
End Sub

Private Sub CreateButton()
    If MenuType = VbeMenu Then
        Set MenuEvent = New CVBECommandHandler
    End If
    Set CmdBarItem = MenuItem.Controls.Add
    With CmdBarItem
        .Caption = Caption
        .BeginGroup = Divider
        .OnAction = Action
        If FaceId <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
    If MenuType = VbeMenu Then
        Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
        EventHandlers.Add MenuEvent, CmdBarItem.Caption
    End If
End Sub

Private Sub DirectButton()
    Dim CmdBarItem As CommandBarControl
    If MenuType = VbeMenu Then
        Set MenuEvent = New CVBECommandHandler
    End If
    Select Case MenuLevel
        Case Is = 1
            Set CmdBarItem = TargetCommandbar.Controls.Add
        Case Is = 2
'            If MenuType = RightClickMenu Then
'                Set CmdBarItem = TargetCommandbar.Controls.Add
'            Else
                Set CmdBarItem = MainMenu.Controls.Add
'            End If
    End Select
    With CmdBarItem
        .Style = msoButtonIconAndCaption
        .Caption = Caption
        .BeginGroup = Divider
        .OnAction = Action
        If FaceId <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
    If MenuType = VbeMenu Then
        Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
        EventHandlers.Add MenuEvent
    End If
End Sub

Private Sub DeleteControlsAndHandlers(ws As Worksheet)
    Select Case LCase(ws.Range(rMENU_TYPE))
        Case "vbemenu"
            MenuType = VbeMenu
        Case "worksheetmenu"
            MenuType = WorksheetMenu
        Case "rightclickmenu"
            MenuType = RightClickMenu
    End Select
    If MenuType = VbeMenu Then
        If BarExists(ws.Range(rC_TAG)) Then
            Application.VBE.CommandBars(ws.Range(rC_TAG).Text).Delete
            Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).Text)
        End If
        Rem
    ElseIf MenuType = WorksheetMenu Then
        Set Ctrl = Application.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).Text)
    ElseIf MenuType = RightClickMenu Then
        If BarExists(ws.Range(rC_TAG).Text) Then
            CommandBars(ws.Range(rC_TAG).Text).Delete
        End If
        Exit Sub
    End If
    On Error Resume Next
    Do
        Ctrl.Delete
        If MenuType = VbeMenu Then
            Rem
            Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).Text)
        Else
            Set Ctrl = Application.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).Text)
        End If
    Loop While Not Ctrl Is Nothing
    On Error GoTo 0
    DeleteHandlersFor ws
End Sub

Private Sub DeleteHandlersFor(ws As Worksheet)
    On Error Resume Next
    markControlType ws
    Dim cell As Range
    Set cell = ws.Cells(2, 6)
    Do Until IsEmpty(cell)
        If cell.Text = "Button" Then
            EventHandlers.Remove cell.Offset(0, -3).Text
        End If
        Set cell = cell.Offset(1)
    Loop
End Sub

Private Sub ListWorksheetBars()
    Dim oCB As CommandBar
    Dim oWK As Worksheet
    Set oWK = ThisWorkbook.Sheets("ListSheetBars")
    oWK.Cells.Clear
    Dim arr As Variant
    arr = Array("Bar name", "Chinese name", "Built-in", "Visible")
    Dim iCol As Integer
    iCol = UBound(arr) + 1
    oWK.Range("a1").Resize(1, iCol) = arr
    oWK.Range("a1").Resize(1, iCol).Cells.Font.Bold = True
    Dim i As Long
    i = 2
    Dim cbVar(300, 3) As Variant
    For Each oCB In Excel.Application.CommandBars
        cbVar(i - 2, 0) = oCB.Name
        cbVar(i - 2, 1) = oCB.NameLocal
        cbVar(i - 2, 2) = oCB.BuiltIn
        cbVar(i - 2, 3) = oCB.Visible
        i = i + 1
    Next
    oWK.Cells(2, 1).Resize(UBound(cbVar, 1) + 1, UBound(cbVar, 2) + 1) = cbVar
    oWK.Columns.AutoFit
End Sub

Private Sub ListVBEBars()
    Dim oCB As CommandBar
    Dim oWK As Worksheet
    Set oWK = ThisWorkbook.Sheets("ListVBEBars")
    oWK.Cells.Clear
    Dim arr As Variant
    arr = Array("Bar name", "Chinese name", "Built-in", "Visible")
    Dim iCol As Integer
    iCol = UBound(arr) + 1
    oWK.Range("a1").Resize(1, iCol) = arr
    oWK.Range("a1").Resize(1, iCol).Cells.Font.Bold = True
    Dim i As Long
    i = 2
    Dim cbVar(300, 3) As Variant
    For Each oCB In Application.VBE.CommandBars
        cbVar(i - 2, 0) = oCB.Name
        cbVar(i - 2, 1) = oCB.NameLocal
        cbVar(i - 2, 2) = oCB.BuiltIn
        cbVar(i - 2, 3) = oCB.Visible
        i = i + 1
    Next
    oWK.Cells(2, 1).Resize(UBound(cbVar, 1) + 1, UBound(cbVar, 2) + 1) = cbVar
    oWK.Columns.AutoFit
End Sub

Private Sub exampleOfControls()
    Dim cbc As CommandBarControl
    Dim cbb As CommandBarButton
    Dim cbcm As CommandBarComboBox
    Dim cbp As CommandBarPopup
    With Application.VBE.CommandBars("CodeArchive")
        Set cbc = .Controls.Add(ID:=3, Temporary:=True)
        Set cbb = .Controls.Add(Temporary:=True)
        cbb.Caption = "A new command"
        cbb.Style = msoButtonCaption
        cbb.OnAction = "NewCommand_OnAction"
        Set cbcm = .Controls.Add(Type:=msoControlComboBox, Temporary:=True)
        cbcm.Caption = "Combo:"
        cbcm.AddItem "list entry 1"
        cbcm.AddItem "list entry 2"
        cbcm.OnAction = "NewCommand_OnAction"
        cbcm.Style = msoComboLabel
        Set cbc = .Controls.Add(Type:=msoControlDropdown, Temporary:=True)
        cbc.Caption = "Dropdown:"
        cbc.AddItem "list entry 1"
        cbc.AddItem "list entry 2"
        cbc.OnAction = "MenuDropdown_OnAction"
        Set cbp = .Controls.Add(Type:=msoControlPopup, Temporary:=True)
        cbp.Caption = "new sub menu"
        Set cbb = cbp.Controls.Add
        cbb.Caption = "sub entry 1"
        Set cbb = cbp.Controls.Add
        cbb.Caption = "sub entry 2"
    End With
End Sub

Private Sub ImageFromEmbedded()
    Dim p As Excel.Picture
    Dim Btn As Office.CommandBarButton
    Set Btn = Application.CommandBars.FindControl(ID:=30007) _
        .Controls.Add(Type:=msoControlButton, Temporary:=True)
    Set p = Worksheets("Sheet1").Pictures("ThePict")
    p.CopyPicture xlScreen, xlBitmap
    With Btn
        .Caption = "Click Me"
        .Style = msoButtonIconAndCaption
        .PasteFace
    End With
End Sub

Private Sub ImageFromExternalFile()
    Dim Btn As Office.CommandBarButton
    Set Btn = Application.CommandBars.FindControl(ID:=30007) _
        .Controls.Add(Type:=msoControlButton, Temporary:=True)
    With Btn
        .Caption = "Click Me"
        .Style = msoButtonIconAndCaption
        .Picture = LoadPicture("C:\TestPic.bmp")
    End With
End Sub

Private Sub ResetCBAR()
    Excel.Application.CommandBars("Cell").Reset
End Sub

Private Sub TestCallingSubFromFile()
    Application.Run "'" & Workbooks("test.xlsm").FullName & "'!Testme"
End Sub



Public Function IsLoaded(formName As String) As Boolean
    Dim Frm As Object
    For Each Frm In VBA.UserForms
        If Frm.Name = formName Then
            IsLoaded = True
            Exit Function
        End If
    Next Frm
    IsLoaded = False
End Function

Sub openUValiationDropdown()
    Dim lngValType As Long
    On Error Resume Next
    lngValType = ActiveCell.Validation.Type
    On Error GoTo 0
    Select Case lngValType
        Case Is = 3
            uValidationDropdown.Show
        Case Else
            If IsLoaded("uValidationDropdown") Then
                On Error Resume Next
                Unload uValidationDropdown
                On Error GoTo 0
            End If
    End Select
End Sub

