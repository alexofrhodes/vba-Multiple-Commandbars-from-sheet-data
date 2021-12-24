Attribute VB_Name = "BarBuilder"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DEV                : ANASTASIOU ALEX
'GIT                 : https://github.com/alexofrhodes
'FEEDBACK   : ANASTASIOUALEX@GMAIL.COM
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PROJECT    : RAISE THE BAR
'LICENSE     : MIT
'PURPOSE    : CREATE COMMAND BARS FROM SHEET DATA
'VERSION    :  2021/07/31  Initial Release
'                        2021/12/25    Take input from sheet data instead of userform
'                                               Create a commandbar from each sheet
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OPTIONS    :
'             > WORKBOOK CONTEXT
'                 -Worksheet Menu Bar (Add-ins tab)
'                 -Cell
'                 -Column
'                 -Row
'             > VBE Context
'                 -Menu Bar
'                    -Existing Menu Bar Controls
'                    -New Control in Menu Bar
'                    -Floating
'                 -Code Window
'                 -Project Window
'                 -Edit Toolbar
'                 -Debug Toolbar
'                 -Userform Toolbar
'             > Independent Popup CommandBar (e.g. to call with right click)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Public Const C_TAG = "MY_VBE_TAG"
'Public BAR_EPIC As String

Public Const rBUILD_ON_OPEN = "I2"
Public Const rC_TAG = "I4"
Public Const rMENU_TYPE = "I5"
Public Const rBAR_LOCATION = "I6"
Public Const rTARGET_CONTROL = "I7"

Public C_TAG As String
Public MenuEvent As CVBECommandHandler
Public CmdBarItem As CommandBarControl
Public EventHandlers As New Collection
Public TargetCommandbar As CommandBar
Public TargetControl As CommandBarControl
Public MainMenu As CommandBarControl
Public MenuItem As CommandBarControl
Public ctrl As Office.CommandBarControl
Public MenuLevel, NextLevel, Caption, Divider, FaceId
Public Action As String
Public MenuSheet As Worksheet
Public Row As Integer
Public MenuType As Long
Public Const WorksheetMenu = 1
Public Const VbeMenu = 2
Public Const RightClickMenu = 3
Public BarLocation As String

Sub NewBar()
Dim wsMain As Worksheet
Set wsMain = ThisWorkbook.Worksheets("BAR_Main")
Dim wsCopy As Worksheet
wsMain.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set wsCopy = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
wsCopy.Name = "BAR_" & lastBar + 1
wsCopy.Range("A1").CurrentRegion.Offset(1).ClearContents
wsCopy.Range("I4:I7").ClearContents
wsCopy.Range("I2") = False
End Sub

Function lastBar() As Long
Dim counter As Long
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
If UCase(Left(ws.Name, 4)) = "BAR_" Then counter = counter + 1
Next
lastBar = counter
End Function

Sub CreateAllBars()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
If UCase(Left(ws.Name, 4)) = "BAR_" Then
If ws.Range(rBUILD_ON_OPEN) = True Then CommandBarBuilder (ws)
End If
Next
End Sub
Sub DeleteAllBars()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
If UCase(Left(ws.Name, 4)) = "BAR_" Then DeleteControlsAndHandlers ws
Next
End Sub

Sub ShowUF2()
    UserForm2.Show vbModeless
End Sub
Function SetCMDbar(ws As Worksheet) As Boolean
    'BAR_EPIC = "vbaCodeArchive"
    
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
        Case Else                                ' FLOATING
            'If BarExists(BAR_EPIC) = False Then
            Set TargetCommandbar = Application.VBE.CommandBars.Add(C_TAG, Position:=msoBarTop, Temporary:=True) 'msoBarFloating
                 TargetCommandbar.Visible = True
        End Select
    ElseIf MenuType = WorksheetMenu Then
        Select Case ws.Range(rBAR_LOCATION)
        Case Is = "Worksheet Menu Bar", "Cell", "Column", "Row"
            Set TargetCommandbar = Application.CommandBars(BarLocation)
            SetCMDbar = True
        Case Else
            '
        End Select
    Else
    End If
End Function
Function BarExists(findBarName As String) As Boolean
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
Sub BuildBarFromShape()
CommandBarBuilder ActiveSheet.Shapes(Application.Caller).Parent
End Sub
Sub DeleteBarFromShape()
DeleteControlsAndHandlers ActiveSheet.Shapes(Application.Caller).Parent
End Sub
Sub CommandBarBuilder(ws As Worksheet)
    DeleteControlsAndHandlers ws
    SetCMDbar ws
    Set MenuSheet = ws
    Row = 2
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
    Do Until IsEmpty(MenuSheet.Cells(Row, 1))
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
        Row = Row + 1
        ReSetVariables
    Loop
    Debug.Print "Bar created"
End Sub
Sub SetVariables()
    With MenuSheet
        MenuLevel = .Cells(Row, 1)
        Caption = .Cells(Row, 2)
        Action = .Cells(Row, 3)
        Divider = .Cells(Row, 4)
        FaceId = .Cells(Row, 5)
        NextLevel = .Cells(Row + 1, 1)
    End With
End Sub
Sub ReSetVariables()
        MenuLevel = ""
        Caption = ""
        Action = ""
        Divider = ""
        FaceId = ""
        NextLevel = ""
End Sub
Sub CreateMainMenu()
    If MenuType = VbeMenu Then
        'Set MainMenu = TargetControl.Controls.Add(Type:=msoControlPopup)
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup)
    ElseIf MenuType = WorksheetMenu Then
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    ElseIf MenuType = RightClickMenu Then
        Set TargetCommandbar = CommandBars.Add(C_TAG, msoBarPopup, , True)
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup)
        Exit Sub
    End If
    With MainMenu
        .Caption = Caption
        .BeginGroup = Divider
        If FaceId <> "" And Action <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
End Sub
Sub CreatePopup()
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
Sub CreateButton()
    If MenuType = VbeMenu Then
        Set MenuEvent = New CVBECommandHandler
    End If
    Set CmdBarItem = MenuItem.Controls.Add
    With CmdBarItem
        .Caption = Caption
        .BeginGroup = Divider
        .OnAction = Action                       '    "'" & ThisWorkbook.Name & "'!" & Action
        If FaceId <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
    If MenuType = VbeMenu Then
        Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
        EventHandlers.Add MenuEvent
    End If
End Sub
Sub DirectButton()
Dim CmdBarItem As CommandBarControl
   If MenuType = VbeMenu Then
        Set MenuEvent = New CVBECommandHandler
    End If
    Select Case MenuLevel
    Case Is = 1
        Set CmdBarItem = TargetCommandbar.Controls.Add
    Case Is = 2
        If MenuType = RightClickMenu Then
            Set CmdBarItem = TargetCommandbar.Controls.Add
        Else
            Set CmdBarItem = MainMenu.Controls.Add
        End If
    End Select
    With CmdBarItem
        .Style = msoButtonIconAndCaption
        .Caption = Caption
        .BeginGroup = Divider
        .OnAction = Action                       '    "'" & ThisWorkbook.Name & "'!" & Action
        If FaceId <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
    If MenuType = VbeMenu Then
        Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
        EventHandlers.Add MenuEvent
    End If
End Sub
Sub DeleteAll()
    MenuType = 1
    DeleteControlsAndHandlers
    MenuType = 2
    DeleteControlsAndHandlers
    MenuType = 3
    DeleteControlsAndHandlers
End Sub
Sub DeleteControlsAndHandlers(ws As Worksheet)

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
            Set ctrl = Application.VBE.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).Text)
        End If
    ElseIf MenuType = WorksheetMenu Then
        Set ctrl = Application.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).Text)
    ElseIf MenuType = RightClickMenu Then
        If BarExists(ws.Range(rC_TAG).Text) Then
            CommandBars(ws.Range(rC_TAG).Text).Delete
        End If
        Exit Sub
    End If
    Do Until ctrl Is Nothing
        ctrl.Delete
        If MenuType = VbeMenu Then
            Set ctrl = Application.VBE.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).Text)
        Else
            Set ctrl = Application.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).Text)
        End If
    Loop
    If MenuType = VbeMenu Then
        Do Until EventHandlers.Count = 0
            EventHandlers.Remove 1
        Loop
    End If
End Sub
Sub ListBars()
ListWorksheetBars
ListVBEBars
End Sub
Sub ListWorksheetBars()
    'Sheets.Add(before:=Sheets(1)).Select
    Dim oCB As CommandBar
    Dim oWK As Worksheet
    Set oWK = ThisWorkbook.Sheets("SheetBars")
    oWK.Cells.Clear
    Dim arr As Variant
    arr = Array("Bar name", "Chinese name", "Built-in", "Visible")
    Dim iCol As Integer
    iCol = UBound(arr) + 1
    oWK.Range("a1").Resize(1, iCol) = arr
    oWK.Range("a1").Resize(1, iCol).Cells.Font.Bold = True
    Dim I As Long
    I = 2
    Dim cbVar(300, 3) As Variant
    For Each oCB In Excel.Application.CommandBars
        cbVar(I - 2, 0) = oCB.Name
        cbVar(I - 2, 1) = oCB.NameLocal
        cbVar(I - 2, 2) = oCB.BuiltIn
        cbVar(I - 2, 3) = oCB.Visible
        I = I + 1
    Next
    oWK.Cells(2, 1).Resize(UBound(cbVar, 1) + 1, UBound(cbVar, 2) + 1) = cbVar
    oWK.Columns.AutoFit
End Sub
Sub ListVBEBars()
    'Sheets.Add(before:=Sheets(1)).Select
    Dim oCB As CommandBar
    Dim oWK As Worksheet
    Set oWK = ThisWorkbook.Sheets("VBEBars")
    oWK.Cells.Clear
    Dim arr As Variant
    arr = Array("Bar name", "Chinese name", "Built-in", "Visible")
    Dim iCol As Integer
    iCol = UBound(arr) + 1
    oWK.Range("a1").Resize(1, iCol) = arr
    oWK.Range("a1").Resize(1, iCol).Cells.Font.Bold = True
    Dim I As Long
    I = 2
    Dim cbVar(300, 3) As Variant
    For Each oCB In Application.VBE.CommandBars
        cbVar(I - 2, 0) = oCB.Name
        cbVar(I - 2, 1) = oCB.NameLocal
        cbVar(I - 2, 2) = oCB.BuiltIn
        cbVar(I - 2, 3) = oCB.Visible
        I = I + 1
    Next
    oWK.Cells(2, 1).Resize(UBound(cbVar, 1) + 1, UBound(cbVar, 2) + 1) = cbVar
    oWK.Columns.AutoFit
End Sub
Sub runListProcs()
ThisWorkbook.Sheets("Procs").Cells.Clear
Dim procArray As Variant
procArray = ListProcs(ThisWorkbook)
ThisWorkbook.Sheets("Procs").Range("A1").Resize(UBound(procArray) + 1) = _
WorksheetFunction.Transpose(procArray)
End Sub
Function ListProcs(wb As Workbook) As Variant
    Dim varr As Variant
    Dim vbComp As VBComponent
    Dim var As Variant
    Dim Output As Collection
    Set Output = New Collection
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type <> vbext_ct_StdModule Then GoTo Skip
            'If vbComp.Name <> "ListMyProcs" Then GoTo Skip
                var = Split(ProcList(vbComp), Chr(10))
                For I = LBound(var) To UBound(var)
                    Output.Add var(I)
                Next I
Skip:
    Next vbComp
    ReDim varr(Output.Count)
    I = 0
    For Each element In Output
        varr(I) = element
        I = I + 1
    Next element
ListProcs = varr
End Function
Public Function ProcList(vbComp As VBComponent) As Variant
Dim codeMod As CodeModule
Set codeMod = vbComp.CodeModule
    Dim LineNum As Long
    Dim NumLines As Long
    Dim ProcName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim out As Variant
    LineNum = codeMod.CountOfDeclarationLines + 1
    Do Until LineNum >= codeMod.CountOfLines
        ProcName = codeMod.ProcOfLine(LineNum, ProcKind)
        If out = vbNullString Then
            out = ProcName
        Else
            out = out & vbNewLine & ProcName
        End If
        LineNum = codeMod.ProcStartLine(ProcName, ProcKind) + codeMod.ProcCountLines(ProcName, ProcKind) + 1
    Loop
    ProcList = out
End Function

'''''''''''''''''
'''''NOTES'''''''
'''''''''''''''''
'----------------------
'WORKSHEET COMMAND BARS
'----------------------
'Application.CommandBars("Worksheet Menu Bar").Controls.Add
'Application.CommandBars("Cell").Controls.Add
'Application.CommandBars("Column").Controls.Add
'Application.CommandBars("Row").Controls.Add
'----------------------
''VBE COMMAND BARS
'----------------------
'----------------------
''add your own command bar
'----------------------
'With Application.VBE.CommandBars.Add("CodeArchive", Position:=msoBarFloating, Temporary:=True)
'    .Visible = True
'End With
'Application.VBE.CommandBars("CodeArchive").Delete
'----------------------
''use existing command bars
'----------------------
'Set TargetControl = Application.VBE.CommandBars("Menu Bar").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Code Window").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Project Window").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Edit").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Debug").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Userform").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'----------------------
''use existing controls
'----------------------
'Set TargetControl = Application.VBE.CommandBars("Menu Bar").Controls.("Tools")
'-----------
'Use combobox
'-----------
''call a sub through class events handler
''the sub to contain the following
'With Application.VBE.ActiveCodePane
'  Text = Application.VBE.CommandBars(mcToolBar).Controls(mcInsertList).Text
'  .GetSelection StartLine, StartColumn, EndLine, EndColumn
'  .CodeModule.InsertLines StartLine, Text
'  .SetSelection StartLine, 1, StartLine, 1
'End With
