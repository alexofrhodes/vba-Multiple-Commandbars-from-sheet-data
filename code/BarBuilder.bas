Attribute VB_Name = "BarBuilder"
'PROJECT    : RAISE THE BAR
'LICENSE    : MIT
'
'DEV        : ANASTASIOU ALEX
'GIT        : https://github.com/alexofrhodes
'FEEDBACK   : ANASTASIOUALEX@GMAIL.COM
'
'PURPOSE    : CREATE COMMAND BARS FROM SHEET DATA
'
'OPTIONS    : > WORKBOOK CONTEXT
'                 -Worksheet Menu Bar (Add-ins tab)
'                 -Cell
'                 -Column
'                 -Row
'             > VBE Context
'                 -Menu Bar
'                    -Existing Menu Bar Controls
'                    -New Control in Menu Bar
'                 -Code Window
'                 -Project Window
'                 -Edit Toolbar
'                 -Debug Toolbar
'                 -Userform Toolbar
'             > Independent Popup CommandBar (e.g. to call with right click)
'
'VERSION    :   2021/07/31  Initial Release
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Public Const C_TAG = "MY_VBE_TAG"
Public C_TAG As String
Public Const SheetName = "RaiseTheBar"

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
Const WorksheetMenu = 1
Const VbeMenu = 2
Const RightClickMenu = 3

Public BarLocation As String

Sub ShowCreateBars()
    CreateBars.Show vbModeless
End Sub

Sub ShowDeleteBars()
    DeleteBars.Show vbModeless
End Sub

Sub ShowUF2()
    UserForm2.Show vbModeless
End Sub

Function SetCMDbar() As Boolean
    C_TAG = CreateBars.BarTag.Text
    
    Select Case LCase(CreateBars.BarList.List(CreateBars.BarList.ListIndex))
    Case Is = LCase("WorksheetMenu")
        MenuType = WorksheetMenu
    Case Is = LCase("vbeMenu")
        MenuType = VbeMenu
    Case Is = LCase("RightClickMenu")
        MenuType = RightClickMenu
    Case Else
    End Select
    
    If CreateBars.BarLocation.ListIndex <> -1 Then
        BarLocation = CreateBars.BarLocation.List(CreateBars.BarLocation.ListIndex)
    Else
        BarLocation = 0
    End If
    
    If MenuType = VbeMenu Then
        Select Case BarLocation
        Case Is = "Menu Bar", "Code Window", "Project Window", "Edit", "Debug", "Userform"
            Set TargetCommandbar = Application.VBE.CommandBars(BarLocation)
            SetCMDbar = True
        Case Else                                ' FLOATING
            Set TargetCommandbar = Application.VBE.CommandBars.Add(C_TAG, Position:=msoBarTop, Temporary:=True) 'msoBarFloating
            
                 TargetCommandbar.Visible = True
        End Select
    ElseIf MenuType = WorksheetMenu Then
        Select Case BarLocation
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
        End If
    Next bar
End Function

Sub CommandBarBuilder()
    SetCMDbar
    Set MenuSheet = ThisWorkbook.Sheets(SheetName)
    Row = 2
    
    DeleteControlsAndHandlers
    
    If MenuType = VbeMenu Then
        If BarLocation = "Menu Bar" Then
            If CreateBars.vbeBarControls.ListIndex = 0 Then
                Set TargetControl = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
            Else
                Dim vbControl As String
                vbControl = CreateBars.vbeBarControls.List(CreateBars.vbeBarControls.ListIndex)
                Set TargetControl = TargetCommandbar.Controls(vbControl).Controls.Add(Type:=msoControlPopup, Temporary:=True)
            End If
        Else
            Set TargetControl = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
        End If
        TargetControl.Caption = C_TAG
        TargetControl.Tag = C_TAG
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
    ThisWorkbook.Sheets("combarTAGS").Range("A99999").End(xlUp).Offset(1, 0).Value = C_TAG
    MsgBox "Bar created"
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
        Set MainMenu = TargetControl.Controls.Add(Type:=msoControlPopup)
    ElseIf MenuType = WorksheetMenu Then
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    ElseIf MenuType = RightClickMenu Then
        Set TargetCommandbar = CommandBars.Add(C_TAG, msoBarPopup, , True)
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

Sub DeleteControlsAndHandlers()
    If MenuType = VbeMenu Then
        Set ctrl = Application.VBE.CommandBars.FindControl(Tag:=C_TAG)
    ElseIf MenuType = WorksheetMenu Then
        Set ctrl = Application.CommandBars.FindControl(Tag:=C_TAG)
    ElseIf MenuType = RightClickMenu Then
        If BarExists(C_TAG) Then
            CommandBars(C_TAG).Delete
        End If
        Exit Sub
    End If
    
    Do Until ctrl Is Nothing
        ctrl.Delete
        If MenuType = VbeMenu Then
            Set ctrl = Application.VBE.CommandBars.FindControl(Tag:=C_TAG)
        Else
            Set ctrl = Application.CommandBars.FindControl(Tag:=C_TAG)
        End If
    Loop
        
    If MenuType = VbeMenu Then
        Do Until EventHandlers.Count = 0
            EventHandlers.Remove 1
        Loop
    End If
End Sub

Sub ListCommandBars()
    'Sheets.Add(before:=Sheets(1)).Select
    Dim oCB As CommandBar
    Dim oWK As Worksheet
    Set oWK = ThisWorkbook.Sheets("combarLIST")
    
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

Sub runListProcs()
ThisWorkbook.Sheets("Procs").Cells.Clear
Dim procArray As Variant
procArray = ListProcs(ThisWorkbook)
ThisWorkbook.Sheets("Procs").Range("A1").Resize(UBound(procArray) + 1) = _
WorksheetFunction.Transpose(procArray)
End Sub
Function ListProcs(wb As Workbook) As Variant
    Dim vArr As Variant
    Dim vbComp As VBComponent
    Dim var As Variant
    Dim output As Collection
    Set output = New Collection
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type <> vbext_ct_StdModule Then GoTo Skip
            If vbComp.Name <> "ListMyProcs" Then GoTo Skip
                var = Split(ProcList(vbComp.CodeModule), Chr(10))
                For i = LBound(var) To UBound(var)
                    output.Add var(i)
                Next i
Skip:
    Next vbComp
    ReDim vArr(output.Count)
    i = 0
    For Each element In output
        vArr(i) = element
        i = i + 1
    Next element
ListProcs = vArr
End Function

Function ProcList(codeMod As CodeModule) As Variant
    Dim lineNum As Long
    Dim NumLines As Long
    Dim procName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    lineNum = codeMod.CountOfDeclarationLines + 1
    Do Until lineNum >= codeMod.CountOfLines
        procName = codeMod.ProcOfLine(lineNum, ProcKind)
        If ProcList = vbNullString Then
            ProcList = procName
        Else
            ProcList = ProcList & Chr(10) & procName
        End If
        lineNum = codeMod.ProcStartLine(procName, ProcKind) + codeMod.ProcCountLines(procName, ProcKind) + 1
    Loop
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


