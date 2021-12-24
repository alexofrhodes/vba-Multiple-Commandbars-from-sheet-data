Attribute VB_Name = "Format"

Public Enum ProcScope
    ScopePrivate = 1
    ScopePublic = 2
    ScopeFriend = 3
    ScopeDefault = 4
End Enum

Public Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
End Enum

Public Type ProcInfo
    procName As String
    ProcKind As VBIDE.vbext_ProcKind
    ProcStartLine As Long
    ProcBodyLine As Long
    ProcCountLines As Long
    ProcScope As ProcScope
    ProcDeclaration As String
End Type

Public Function FormatHelp() As String

Dim Q As String: Q = "'"
Dim QQ As String: QQ = """"
Dim s As String
s = s & vbNewLine & "DEV Anastasiou Alex - https://github.com/alexofrhodes - anastasioualex@gmail.com"
s = s & vbNewLine & "These procedures modify the code. Be sure to test and confirm how they work before using on important files. Always keep backups."
s = s & vbNewLine & "Their scope is either the codepane selection, the active procedure (where the mouse is inside of), active module (whose codepane the mouse is inside of),"
s = s & vbNewLine & "the active workbook (whose module's codepane the mouse is inside of), or a parameter you  must pass (procedure name, vbComponent or workbook)"
s = s & vbNewLine & "If vbe is stopped then the onAction of vbe commandbars stop working. Use the following line of code from the immediate window to restore:"
s = s & vbNewLine & "Application.Run '" & ThisWorkbook.FullName & "'!CommandBarBuilder" & vbNewLine

s = s & vbNewLine & "ActiveProcButton"
s = s & vbNewLine & "ArguementsForCalling"
s = s & vbNewLine & "ArguementsForDefining"
s = s & vbNewLine & "BlankLinesRemoveModule"
s = s & vbNewLine & "BlankLinesRemoveProcedure"
s = s & vbNewLine & "BlankLinesRemoveWorkbook"
s = s & vbNewLine & "CaseLower"
s = s & vbNewLine & "CaseProper"
s = s & vbNewLine & "CaseUpper"
s = s & vbNewLine & "CommentsRemoveModule"
s = s & vbNewLine & "CommentsRemoveProcedure"
s = s & vbNewLine & "CommentsRemoveWorkbook"
s = s & vbNewLine & "DebugDisableModule"
s = s & vbNewLine & "DebugDisableProcedure"
s = s & vbNewLine & "DebugDisableWorkbook"
s = s & vbNewLine & "DebugEnableModule"
s = s & vbNewLine & "DebugEnableProcedure"
s = s & vbNewLine & "DebugEnableWorkbook"
s = s & vbNewLine & "Encapsulate " & QQ & "(" & QQ & "," & QQ & ")" & QQ
s = s & vbNewLine & "EncapsulateMultiple " & "vbNewLine" & "," & QQ & "(" & QQ & "," & QQ & ")" & QQ
s = s & vbNewLine & "Flip " & QQ & "=" & QQ
s = s & vbNewLine & "FlipMultiple " & "vbnewline," & QQ & "=" & QQ
s = s & vbNewLine & "IndentModule"
s = s & vbNewLine & "IndentWorkbook"
s = s & vbNewLine & "Inject " & QQ & "You can type here or use a function which returns a string" & QQ
s = s & vbNewLine & "NumberLinesModuleAdd"
s = s & vbNewLine & "NumberLinesModuleRemove"
s = s & vbNewLine & "NumberLinesProcedureAdd"
s = s & vbNewLine & "NumberLinesProcedureRemove"
s = s & vbNewLine & "SetToNothing"
s = s & vbNewLine & "SortProceduresModule"
s = s & vbNewLine & "SortProceduresWorkbook"
s = s & vbNewLine & "SubstituteInSelection"
s = s & vbNewLine & "sortSelection" & " vbnewline"

FormatHelp = s
Debug.Print FormatHelp
End Function

''''
Sub ActiveProcButton()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select cell to contain the shape"
        Exit Sub
    End If
    With addShape
        '.name="ProcButton_" & activeprocname
        .OnAction = ActiveProcName
        .TextFrame2.TextRange.Text = ActiveProcName
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame2.WordWrap = msoFalse
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .Left = Selection.Left
        .Top = Selection.Top
    End With
End Sub

Function addShape() As Shape
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes.addShape _
              (msoShapeRoundedRectangle, 1, 1, 500, 10)
    With shp.ThreeD
        .BevelTopType = msoBevelCircle
        .BevelTopInset = 6
        .BevelTopDepth = 6
    End With

    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
        .Solid
    End With
    With shp.line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With

    Set addShape = shp
End Function


'''''
Function ProceduresOfWorkbook(wb As Workbook) As Collection
    Dim vbComp As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim LineNum As Long
    Dim Coll As New Collection
    Dim procName As String
    For Each vbComp In wb.VBProject.VBComponents
        With vbComp.CodeModule
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines
                procName = .ProcOfLine(LineNum, ProcKind)
                Coll.Add procName
                LineNum = .ProcStartLine(procName, ProcKind) + .ProcCountLines(procName, ProcKind) + 1
            Loop
        End With
    Next
    Set ProceduresOfWorkbook = Coll
End Function
Function ProceduresOfModule(vbComp As VBComponent) As Collection
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim LineNum As Long
    Dim Coll As New Collection
    Dim procName As String
        With vbComp.CodeModule
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines
                procName = .ProcOfLine(LineNum, ProcKind)
                Coll.Add procName
                LineNum = .ProcStartLine(procName, ProcKind) + .ProcCountLines(procName, ProcKind) + 1
            Loop
        End With
    Set ProceduresOfModule = Coll
End Function
Function ProcedureFirstLine(vbComp As VBComponent, procName As String) As Long
    Dim codeMod As CodeModule
    Set codeMod = vbComp.CodeModule
    Dim n As Long
    Dim s As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    With codeMod
        For n = .ProcBodyLine(procName, ProcKind) To .CountOfLines
            s = .Lines(n, 1)
            If Trim(s) = vbNullString Then
                ' blank line, skip it
            ElseIf Left(Trim(s), 1) = "'" Then
                ' comment line, skip it
            ElseIf Right(Trim(s), 1) = "_" Then
                'skip
            ElseIf Right(Trim(s), 1) = ")" Then
                'skip
            ElseIf InStrExact(1, s, "Sub ") Or InStrExact(1, s, "Function ") Then
                'skip
            Else
                Exit For
            End If
        Next n
    End With
    ProcedureFirstLine = n
End Function
Public Function GetProcText(vbComp As VBComponent, _
                            sProcName As String, _
                            Optional bInclHeader As Boolean = True)
    '#IMPORTS getProcKind
    Dim codeMod As CodeModule
    Set codeMod = vbComp.CodeModule
    Dim lProcStart            As Long
    Dim lProcBodyStart        As Long
    Dim lProcNoLines          As Long
    Const vbext_pk_Proc = 0
    On Error GoTo Error_Handler
    lProcStart = codeMod.ProcStartLine(sProcName, getProcKind(vbComp, sProcName))
    lProcBodyStart = codeMod.ProcBodyLine(sProcName, getProcKind(vbComp, sProcName))
    lProcNoLines = codeMod.ProcCountLines(sProcName, getProcKind(vbComp, sProcName))
    If bInclHeader = True Then
        GetProcText = codeMod.Lines(lProcStart, lProcNoLines)
    Else
        lProcNoLines = lProcNoLines - (lProcBodyStart - lProcStart)
        GetProcText = codeMod.Lines(lProcBodyStart, lProcNoLines)
    End If
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    Debug.Print "The following error has occurred" & vbCrLf & vbCrLf & _
                "Error Number: " & Err.Number & vbCrLf & _
                "Error Source: GetProcText" & vbCrLf & _
                "Error Description: " & Err.description & _
                Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Line No: " & Erl)
    Resume Error_Handler_Exit
End Function
Function ModuleOfProcedure(wb As Workbook, ProcedureName As String) As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim LineNum As Long, NumProc As Long
    Dim procName As String
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        With vbComp.CodeModule
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines
                procName = .ProcOfLine(LineNum, ProcKind)
                LineNum = .ProcStartLine(procName, ProcKind) + .ProcCountLines(procName, ProcKind) + 1
                If UCase(procName) = UCase(ProcedureName) Then
                    Set ModuleOfProcedure = vbComp
                    Exit Function
                End If
            Loop
        End With
    Next vbComp
End Function
Public Function InStrExact(Start As Long, SourceText As String, WordToFind As String, _
                    Optional CaseSensitive As Boolean = False, _
                    Optional AllowAccentedCharacters As Boolean = False) As Long
    Dim x As Long, str1 As String, str2 As String, Pattern As String
    Const UpperAccentsOnly As String = "Ã‡Ã‰Ã‘"
    Const UpperAndLowerAccents As String = "Ã‡Ã‰Ã‘Ã§Ã©Ã±"
    If CaseSensitive Then
        str1 = SourceText
        str2 = WordToFind
        Pattern = "[!A-Za-z0-9]"
        If AllowAccentedCharacters Then Pattern = Replace(Pattern, "!", "!" & UpperAndLowerAccents)
    Else
        str1 = UCase(SourceText)
        str2 = UCase(WordToFind)
        Pattern = "[!A-Z0-9]"
        If AllowAccentedCharacters Then Pattern = Replace(Pattern, "!", "!" & UpperAccentsOnly)
    End If
    For x = Start To Len(str1) - Len(str2) + 1
        If Mid(" " & str1 & " ", x, Len(str2) + 2) Like Pattern & str2 & Pattern _
                                                   And Not Mid(str1, x) Like str2 & "'[" & Mid(Pattern, 3) & "*" Then
            InStrExact = x
            Exit Function
        End If
    Next
End Function

Public Function getProcKind(vbComp As VBComponent, ByVal sProcName As String) As Long
'#IMPORTS InStrExact
    Dim codeMode As CodeModule
    Set codeMode = vbComp.CodeModule
    Const vbext_pk_Proc As Long = 0
    Const vbext_pk_Let As Long = 1
    Const vbext_pk_Set As Long = 2
    Const vbext_pk_Get As Long = 3
    Dim Txt As String
    Txt = codeMode.Lines(1, codeMode.CountOfLines)
    If InStrExact(1, Txt, "Get " & sProcName) > 0 Then
        getProcKind = 3
    ElseIf InStrExact(1, Txt, "Let " & sProcName) > 0 Then
        getProcKind = 1
    ElseIf InStrExact(1, Txt, "Set " & sProcName) > 0 And Not (InStrExact(1, Txt, "Sub " & sProcName) > 0 Or InStrExact(1, Txt, "Function " & sProcName) > 0) Then
        getProcKind = 2
    Else
        getProcKind = 0
    End If
End Function

Public Function CLIP(Optional StoreText As String) As String
    Dim x As Variant
    x = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
            Case Len(StoreText)
                .SetData "text", x
            Case Else
                CLIP = .GetData("text")
            End Select
        End With
    End With
End Function


Public Function ActiveProcName() As String
    Application.VBE.ActiveCodePane.GetSelection L1&, C1&, L2&, C2&
    ActiveProcName = Application.VBE.ActiveCodePane _
                     .CodeModule.ProcOfLine(L1&, vbext_pk_Proc)
End Function

Public Function ActiveCodepaneWorkbook() As Workbook
    Dim TmpStr As String
    TmpStr = Application.VBE.SelectedVBComponent.Collection.Parent.Filename
    TmpStr = Right(TmpStr, Len(TmpStr) - InStrRev(TmpStr, "\"))
    Set ActiveCodepaneWorkbook = Workbooks(TmpStr)
End Function

Public Function ActiveComp() As VBComponent
    Set ActiveComp = Application.VBE.SelectedVBComponent
End Function

Public Function CommentsRemoveWorkbook()
'#IMPORTS ActiveCodepaneWorkbook
    Dim n               As Long
    Dim I               As Long
    Dim J               As Long
    Dim K               As Long
    Dim l               As Long
    Dim LineText        As String
    Dim exitString      As String
    Dim Quotes          As Long
    Dim Q               As Long
    Dim StartPos        As Long
    Dim vbComp As VBIDE.VBComponent
    For Each vbComp In ActiveCodepaneWorkbook.VBProject.VBComponents
        With vbComp.CodeModule
            For J = .CountOfLines To 1 Step -1
                LineText = Trim(.Lines(J, 1))
                If LineText = "ExitString = " & _
                   """" & "Ignore Comments In This Module" & """" Then
                    Exit For
                End If
                StartPos = 1
Retry:
                n = InStr(StartPos, LineText, "'")
                Q = InStr(StartPos, LineText, """")
                Quotes = 0
                If Q < n Then
                    For l = 1 To n
                        If Mid(LineText, l, 1) = """" Then
                            Quotes = Quotes + 1
                        End If
                    Next l
                End If
                If Quotes = Application.WorksheetFunction.Odd(Quotes) Then
                    StartPos = n + 1
                    GoTo Retry:
                Else
                    Select Case n
                    Case Is = 0
                    Case Is = 1
                        .DeleteLines J, 1
                    Case Is > 1
                        .REPLACELINE J, Left(LineText, n - 1)
                    End Select
                End If
            Next J
        End With
    Next vbComp
    exitString = "Ignore Comments In This Module"
End Function

Public Sub CommentsRemoveModule()
'#IMPORTS ActiveComp
    Dim vbComp As VBComponent
    Set vbComp = ActiveComp
    Dim n               As Long
    Dim I               As Long
    Dim J               As Long
    Dim K               As Long
    Dim l               As Long
    Dim LineText        As String
    Dim exitString      As String
    Dim Quotes          As Long
    Dim Q               As Long
    Dim StartPos        As Long
    With vbComp.CodeModule
        For J = .CountOfLines To 1 Step -1
            LineText = Trim(.Lines(J, 1))
            If LineText = "ExitString = " & _
               """" & "Ignore Comments In This Module" & """" Then
                Exit For
            End If
            StartPos = 1
Retry:
            n = InStr(StartPos, LineText, "'")
            Q = InStr(StartPos, LineText, """")
            Quotes = 0
            If Q < n Then
                For l = 1 To n
                    If Mid(LineText, l, 1) = """" Then
                        Quotes = Quotes + 1
                    End If
                Next l
            End If
            If Quotes = Application.WorksheetFunction.Odd(Quotes) Then
                StartPos = n + 1
                GoTo Retry:
            Else
                Select Case n
                Case Is = 0
                Case Is = 1
                    .DeleteLines J, 1
                Case Is > 1
                    .REPLACELINE J, Left(LineText, n - 1)
                End Select
            End If
        Next J
    End With
    exitString = "Ignore Comments In This Module"
End Sub

Public Sub CommentsRemoveProcedure()
'#IMPORTS ActiveComp
'#IMPORTS ProcedureEndLine
'#IMPORTS ProcedureStartLine
'#IMPORTS ActiveProcName
    Dim vbComp As VBComponent
    Set vbComp = ActiveComp
    Dim n               As Long
    Dim I               As Long
    Dim J               As Long
    Dim K               As Long
    Dim l               As Long
    Dim LineText        As String
    Dim exitString      As String
    Dim Quotes          As Long
    Dim Q               As Long
    Dim StartPos        As Long
    Dim startLine As Long
    startLine = ProcedureStartLine(vbComp.CodeModule, ActiveProcName)
    Dim EndLine As Long
    EndLine = ProcedureEndLine(vbComp.CodeModule, ActiveProcName)
    With vbComp.CodeModule
        For J = EndLine To startLine Step -1
            LineText = Trim(.Lines(J, 1))
            If LineText = "ExitString = " & _
               """" & "Ignore Comments In This Module" & """" Then
                Exit For
            End If
            StartPos = 1
Retry:
            n = InStr(StartPos, LineText, "'")
            Q = InStr(StartPos, LineText, """")
            Quotes = 0
            If Q < n Then
                For l = 1 To n
                    If Mid(LineText, l, 1) = """" Then
                        Quotes = Quotes + 1
                    End If
                Next l
            End If
            If Quotes = Application.WorksheetFunction.Odd(Quotes) Then
                StartPos = n + 1
                GoTo Retry:
            Else
                Select Case n
                Case Is = 0
                Case Is = 1
                    .DeleteLines J, 1
                Case Is > 1
                    .REPLACELINE J, Left(LineText, n - 1)
                End Select
            End If
        Next J
    End With
    exitString = "Ignore Comments In This Module"
End Sub

Public Function BlankLinesRemoveWorkbook()
'#IMPORTS ActiveCodepaneWorkbook
    Dim I As Long
    For I = 1 To ActiveCodepaneWorkbook.VBProject.VBComponents.Count
        Dim n As Long
        Dim s As String
        Dim LineCount As Long
        Dim vbComp As VBIDE.VBComponent
        Set vbComp = ActiveCodepaneWorkbook.VBProject.VBComponents(I)
        With vbComp.CodeModule
            For n = .CountOfLines To 1 Step -1
                s = .Lines(n, 1)
                If Trim(s) = vbNullString Then
                    .DeleteLines n
                ElseIf Left(Trim(s), 1) = "'" Then
                Else
                    LineCount = LineCount + 1
                End If
            Next n
        End With
    Next I
End Function

Public Sub BlankLinesRemoveModule()
'#IMPORTS ActiveComp
    Dim n As Long
    Dim s As String
    Dim LineCount As Long
    Dim vbComp As VBIDE.VBComponent
    Set vbComp = ActiveComp
    With vbComp.CodeModule
        For n = .CountOfLines To 1 Step -1
            s = .Lines(n, 1)
            If Trim(s) = vbNullString Then
                .DeleteLines n
            ElseIf Left(Trim(s), 1) = "'" Then
            Else
                LineCount = LineCount + 1
            End If
        Next n
    End With
End Sub

Public Sub BlankLinesRemoveProcedure()
'#IMPORTS ActiveComp
'#IMPORTS ProcedureEndLine
'#IMPORTS ProcedureStartLine
'#IMPORTS ActiveProcName
    Dim n As Long
    Dim s As String
    Dim vbComp As VBIDE.VBComponent
    Set vbComp = ActiveComp
    With vbComp.CodeModule
        For n = ProcedureEndLine(vbComp.CodeModule, ActiveProcName) To ProcedureStartLine(vbComp.CodeModule, ActiveProcName) Step -1
            s = .Lines(n, 1)
            If Trim(s) = vbNullString Then
                .DeleteLines n
            ElseIf Left(Trim(s), 1) = "'" Then
            Else
            End If
        Next n
    End With
End Sub

Public Sub SortProceduresWorkbook(Optional wb As Workbook)
'#IMPORTS SortProceduresModule
'#IMPORTS ActiveCodepaneWorkbook
If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
Dim vbComp As VBComponent
For Each vbComp In wb.VBProject.VBComponents
SortProceduresModule vbComp
Next
End Sub
Public Sub SortProceduresModule(vbComp As VBComponent)
'#IMPORTS ProcList
'#IMPORTS ActiveComp
'#IMPORTS SortArray
'#IMPORTS GetProcText
    If vbComp Is Nothing Then Set vbComp = ActiveComp
    Dim vProceduresList As Variant
    Dim procedure As Variant
    Dim varr
    Dim I As Long
    Dim ReplacedProcedures As String
    ReplacedProcedures = ""
    Dim startLine As Long
    Dim TotalLines As Long
    If vbComp.CodeModule.CountOfLines = 0 Then Exit Sub
    varr = Split(ProcList(vbComp.CodeModule), vbNewLine)
    startLine = vbComp.CodeModule.ProcStartLine(varr(0), vbext_pk_Proc)
    TotalLines = vbComp.CodeModule.CountOfLines - vbComp.CodeModule.CountOfDeclarationLines
    Call SortArray(varr, LBound(varr), UBound(varr))
    For I = LBound(varr) To UBound(varr)
        If ReplacedProcedures = "" Then
            ReplacedProcedures = GetProcText(vbComp, CStr(varr(I)))
        Else
            ReplacedProcedures = ReplacedProcedures & vbNewLine & _
                                 GetProcText(vbComp, CStr(varr(I)))
        End If
    Next I
    vbComp.CodeModule.DeleteLines startLine, TotalLines
    vbComp.CodeModule.AddFromString ReplacedProcedures
End Sub

Public Sub SortArray(vArray As Variant, inLow As Long, inHi As Long)
    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    tmpLow = inLow
    tmpHi = inHi
    pivot = vArray((inLow + inHi) \ 2)
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    If (inLow < tmpHi) Then SortArray vArray, inLow, tmpHi
    If (tmpLow < inHi) Then SortArray vArray, tmpLow, inHi
End Sub

'''

Public Sub SetToNothing()
'optional procedure as string = activeProcName,
'if procedure="" then procedure=activeprocname
'set vbcomp = ModuleOfProcedure(procedure)

'#IMPORTS ActiveComp
'#IMPORTS ProcedureEndLine
'#IMPORTS ProcedureFirstLine
'#IMPORTS ActiveProcName
    Dim vbComp As VBComponent
    Set vbComp = ActiveComp
    Dim proc As String
    proc = ActiveProcName
    Dim FIRSTLINE As Long
    Dim LASTLINE As Long
    Dim LINENUMBER As Long
    Dim APPEND As String
    Dim strLine As String
    Dim TERMINATE As String
    With vbComp.CodeModule
        FIRSTLINE = ProcedureFirstLine(vbComp, proc)
        LASTLINE = ProcedureEndLine(vbComp.CodeModule, proc)
        For LINENUMBER = FIRSTLINE To LASTLINE
            strLine = Trim(vbComp.CodeModule.Lines(LINENUMBER, 1))
            If Trim(strLine) Like "Set * = *" Or Trim(strLine) Like "Dim*As New*" Then
                TERMINATE = Split(strLine, " ")(1)
                APPEND = APPEND & vbNewLine & "Set " & TERMINATE & " = Nothing"
            End If
        Next
    End With
    If APPEND <> "" Then vbComp.CodeModule.InsertLines LASTLINE, APPEND
End Sub

Public Function ProcedureEndLine(codeMod As CodeModule, procName As String) As Long
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim StartAt As Long
    Dim EndAt As Long
    Dim CountOf As Long
    StartAt = codeMod.ProcStartLine(procName, ProcKind)
    EndAt = codeMod.ProcStartLine(procName, ProcKind) + codeMod.ProcCountLines(procName, ProcKind) - 1
    CountOf = codeMod.ProcCountLines(procName, ProcKind)
    ProcedureEndLine = EndAt
End Function
Public Function ProcedureStartLine(codeMod As CodeModule, ProcedureName As String) As Long
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim StartAt As Long
    Dim EndAt As Long
    Dim CountOf As Long
   StartAt = codeMod.ProcStartLine(ProcedureName, ProcKind)
   EndAt = codeMod.ProcStartLine(ProcedureName, ProcKind) + codeMod.ProcCountLines(ProcedureName, ProcKind) - 1
   CountOf = codeMod.ProcCountLines(ProcedureName, ProcKind)
   ProcedureStartLine = StartAt
End Function

'''
Public Sub NumberLinesModuleAdd(Optional vbComp As VBComponent)
'#IMPORTS ProceduresOfModule
If vbComp Is Nothing Then Set vbComp = ActiveComp
Dim element
For Each element In ProceduresOfModule(vbComp)
NumberLinesProcedureAdd CStr(element)
Next
End Sub
Public Function NumberLinesProcedureAdd(Optional vbComp As VBComponent, Optional ProcedureName As String)
'#IMPORTS NumberThisLine
'#IMPORTS GetProcText
If vbComp Is Nothing Then Set vbComp = ActiveComp
If ProcedureName = "" Then ProcedureName = ActiveProcName
Dim codeMod As CodeModule
Set codeMod = vbComp.CodeModule
    Dim Txt
    Dim varr
    Dim I As Long
    Dim a As Long
    a = 1
    varr = Split(GetProcText(vbComp, ProcedureName), vbNewLine)
    For I = LBound(varr) To UBound(varr)
        If Txt = "" Then
            If NumberThisLine(varr(I)) Then
                Txt = a & ":" & varr(I)
                a = a + 1
            Else
                Txt = varr(I)
            End If
        Else
            If NumberThisLine(varr(I)) And Right(Trim(varr(I - 1)), 1) <> "_" Then
                Txt = Txt & vbNewLine & a & ":" & varr(I)
                a = a + 1
            Else
                Txt = Txt & vbNewLine & varr(I)
            End If
        End If
    Next I
    startLine = codeMod.ProcStartLine(ProcedureName, vbext_pk_Proc)
    NumLines = codeMod.ProcCountLines(ProcedureName, vbext_pk_Proc)
    codeMod.DeleteLines startLine:=startLine, Count:=NumLines
    codeMod.InsertLines startLine, Txt
End Function

Public Function NumberThisLine(ByVal str As String) As Boolean
    Dim Test As String
    Test = Trim(str)
    If Len(Test) = 0 Then Exit Function
    If Right(Test, 1) = ":" Then Exit Function
    If IsNumeric(Left(Test, 1)) Then Exit Function
    If Test Like "'*" Then Exit Function
    If Test Like "Rem*" Then Exit Function
    If Test Like "Dim*" Then Exit Function
    If Test Like "Sub*" Then Exit Function
    If Test Like "Public*" Then Exit Function
    If Test Like "public*" Then Exit Function
    If Test Like "Function*" Then Exit Function
    If Test Like "End Sub*" Then Exit Function
    If Test Like "End Function*" Then Exit Function
    If Test Like "Debug*" Then Exit Function
    NumberThisLine = True
End Function

Public Sub NumberLinesModuleRemove(Optional vbComp As VBComponent)
'#IMPORTS NumberLinesProcedureRemove
'#IMPORTS ActiveComp
'#IMPORTS ProceduresOfModule

If IsMissing(vbComp) Then Set vbComp = ActiveComp
Dim element
For Each element In ProceduresOfModule(vbComp)
NumberLinesProcedureRemove CStr(element)
Next
End Sub
Public Function NumberLinesProcedureRemove(Optional vbComp As VBComponent, Optional ProcedureName As String)
'#IMPORTS ActiveProcName
'#IMPORTS ActiveComp
'#IMPORTS FirstDigit
'#IMPORTS GetProcText
If vbComp Is Nothing Then Set vbComp = ActiveComp
If ProcedureName = "" Then ProcedureName = ActiveProcName
    Dim codeMod As CodeModule
   Set codeMod = vbComp.CodeModule
    
    Dim startLine As Long
    Dim NumLines As Long
    Dim Txt
    Dim varr
   varr = Split(GetProcText(vbComp, ProcedureName), vbNewLine)
    Dim I As Long
   For I = LBound(varr) To UBound(varr)
       If Txt = "" Then
           If Not IsNumeric(Left(Trim(varr(I)), 1)) Then
               Txt = varr(I)
           Else
               Txt = Left(varr(I), FirstDigit(varr(I)) - 1) & Right(varr(I), Len(varr(I)) - InStr(1, varr(I), ":") - 1)
           End If
       Else
           If Not IsNumeric(Left(Trim(varr(I)), 1)) Then
               Txt = Txt & vbNewLine & varr(I)
           Else
               varr(I) = varr(I) & " "
               Txt = Txt & vbNewLine & Left(varr(I), FirstDigit(varr(I)) - 1) & Right(varr(I), Len(varr(I)) - InStr(1, varr(I), ":") - 1)
           End If
       End If
   Next I
   startLine = codeMod.ProcStartLine(ProcedureName, vbext_pk_Proc)
   NumLines = codeMod.ProcCountLines(ProcedureName, vbext_pk_Proc)
   codeMod.DeleteLines startLine:=startLine, Count:=NumLines
   codeMod.InsertLines startLine, Txt
End Function

Public Function FirstDigit(ByVal strData As String) As Integer
    Dim Re As Object
    Dim REMatches As Object
    Set Re = CreateObject("vbscript.regexp")
    Re.Pattern = "[0-9]"
    Set REMatches = Re.Execute(strData)
    FirstDigit = REMatches(0).FirstIndex + 1
End Function


'''
Public Function IndentWorkbook(Optional wb As Workbook)
'#IMPORTS ActiveCodepaneWorkbook
'#IMPORTS IndentModule
If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
Dim vbComp As VBComponent
For Each vbComp In wb.VBProject.VBComponents
IndentModule vbComp
Next
End Function

Public Function IndentModule(Optional vbComp As VBComponent)
'#IMPORTS ActiveComp
'#IMPORTS IsBlockStart
'#IMPORTS IsBlockEnd
    If vbComp Is Nothing Then Set vbComp = ActiveComp
    Dim codMod As CodeModule
    Set codMod = vbComp.CodeModule
    Dim nIndent As Integer
    Dim nLine As Long
    Dim strNewLine As String
    For nLine = 1 To mCode.CountOfLines
        strNewLine = codMod.Lines(nLine, 1)
        strNewLine = LTrim$(strNewLine)
        If IsBlockEnd(strNewLine) Then nIndent = nIndent - 1
        If nIndent < 0 Then nIndent = 0
        mCode.REPLACELINE nLine, Space$(nIndent * 4) & strNewLine
        If IsBlockStart(strNewLine) Then nIndent = nIndent + 1
    Next nLine
End Function

Public Function IsBlockStart(strLine As String) As Boolean
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = Left$(strLine, nPos)
    Select Case strTemp
    Case "With", "For", "Do", "While", "Select", "Case", "Else", "Else:", "#Else", "#Else:", "Sub", "Function", "Property", "Enum", "Type"
        bOK = True
    Case "If", "#If", "ElseIf", "#ElseIf"
        bOK = (Len(strLine) = (InStr(1, strLine, " Then") + 4))
    Case "public", "Public", "Friend"
        nPos = InStr(1, strLine, " Static ")
        If nPos Then
            nPos = InStr(nPos + 7, strLine, " ")
        Else
            nPos = InStr(Len(strTemp) + 1, strLine, " ")
        End If
        Select Case Mid$(strLine, nPos + 1, InStr(nPos + 1, strLine, " ") - nPos - 1)
        Case "Sub", "Function", "Property", "Enum", "Type"
            bOK = True
        End Select
    End Select
    IsBlockStart = bOK
End Function

Public Function IsBlockEnd(strLine As String) As Boolean
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = Left$(strLine, nPos)
    Select Case strTemp
    Case "Next", "Loop", "Wend", "End Select", "Case", "Else", "#Else", "Else:", "#Else:", "ElseIf", "#ElseIf", "End If", "#End If"
        bOK = True
    Case "End"
        bOK = (Len(strLine) > 3)
    End Select
    IsBlockEnd = bOK
End Function

'''

Public Function ArguementsForCalling() As String
'#IMPORTS ActiveCodepaneWorkbook
'#IMPORTS getArgs
'#IMPORTS CodepaneSelection
'#IMPORTS ModuleOfProcedure
'#IMPORTS CLIP
Dim strip As Boolean: strip = True
    Dim proc As String: proc = CodepaneSelection
    Dim str As String
    str = getArgs(ModuleOfProcedure(ActiveCodepaneWorkbook, proc), proc, strip)
    ArguementsForCalling = str
    CLIP str
    Debug.Print vbNewLine & "Arguement  for calling " & proc & " copied to clipboard" & vbNewLine
    Debug.Print "If it's red it means you must either" & vbNewLine & _
                        "add the word CALL or RUN before it" & vbNewLine & _
                        "Or use a METHOD with it"
End Function
Public Function ArguementsForDefining() As String
'#IMPORTS ActiveCodepaneWorkbook
'#IMPORTS getArgs
'#IMPORTS CodepaneSelection
'#IMPORTS ModuleOfProcedure
'#IMPORTS CLIP
    Dim strip As Boolean: strip = False
    Dim proc As String: proc = CodepaneSelection
    Dim str As String
    str = getArgs(ModuleOfProcedure(proc, ActiveCodepaneWorkbook), proc, strip)
    If str Like "Sub*" Then
        str = str & vbNewLine & vbNewLine & "End Sub"
    Else
        str = str & vbNewLine & vbNewLine & "End Function"
    End If
    ArguementsForDefining = str
    CLIP str
    Debug.Print vbNewLine & "Arguement  for defining " & proc & " copied to clipboard" & vbNewLine
End Function
Public Function getArgsFromClipboard() As String
'#IMPORTS ActiveCodepaneWorkbook
'#IMPORTS getArgs
'#IMPORTS ModuleOfProcedure
    Dim strip As Boolean:   strip = True
    Dim proc As String:     proc = CLIPBOARD
    getArgsFromClipboard = getArgs(ModuleOfProcedure(proc, ActiveCodepaneWorkbook), proc, strip)
End Function

Public Function getArgs(vbComp As VBComponent, procName As String, Optional strip As Boolean) As String
'#IMPORTS modReplaceMulti
'#IMPORTS GetProcedureDeclaration
'#IMPORTS getProcKind
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim str         As Variant
    Dim element     As Long
    Dim line        As String
    Dim firstPart   As String
    Dim secondPart  As String
    Dim Output      As String
    str = GetProcedureDeclaration(vbComp, procName, getProcKind(vbComp, procName))
    str = Replace(str, vbNewLine, "")
    If IsEmpty(str) Then getArgs = "": Exit Function
    If strip = True Then Output = procName & "( _"
    If strip = False Then Output = Left(str, InStr(1, str, "(") - 1) & "( _"
    str = Right(str, Len(str) - InStr(1, str, "("))
    str = Left(str, InStrRev(str, ")") - 1)
    str = Split(str, ",")
    If UBound(str) = -1 Then Exit Function
    For element = LBound(str) To UBound(str)
        If strip = True Then
            line = modReplaceMulti(vbTextCompare, Trim(str(element)), "", "Optional ", "As ", "ByVal ", "ByRef", "ParamArray ", "_")
            firstPart = Split(line, " ")(0)
            secondPart = Split(line, " ")(1)
            Output = Output & vbNewLine & vbTab & IIf(InStr(1, Output, vbNewLine) > 1, ",", "") & firstPart & ":= " & "as" & Split(line, " ")(1) & IIf(element <> UBound(str), " _", ")")
        Else
            line = modReplaceMulti(vbTextCompare, Trim(str(element)), "", "_")
            Output = Output & vbNewLine & vbTab & IIf(InStr(1, Output, vbNewLine) > 1, ",", "") & line & IIf(element <> UBound(str), " _", ")")
        End If
    Next
    getArgs = Output
End Function

Public Function PadLeft(ByVal str As String, ByVal length As Long)
    If Len(str) < length Then
        PadLeft = str + String$(length - Len(str), " ")
    Else
        PadLeft = Left$(str, length)
    End If
End Function

Public Function modReplaceMulti(ByVal Compare As VbCompareMethod, ByVal str As String, toStr As String, _
                                ParamArray replacements() As Variant) As String
    Dim element As Variant
    For Each element In replacements
        str = Replace(str, element, toStr, , , Compare)
    Next
    modReplaceMulti = str
End Function

'''

Public Function DebugEnableWorkbook(Optional wb As Workbook)
'#IMPORTS ActiveCodepaneWorkbook
If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
Dim vbComp As VBComponent
For Each vbComp In wb.VBProject.VBComponents
DebugEnableModule vbComp
Next
End Function
Public Function DebugDisableWorkbook(Optional wb As Workbook)
'#IMPORTS ActiveCodepaneWorkbook
If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
Dim vbComp As VBComponent
For Each vbComp In wb.VBProject.VBComponents
DebugDisableModule vbComp
Next
End Function

Public Sub DebugDisableModule(Optional vbComp As VBComponent)
'#IMPORTS ActiveComp
If vbComp Is Nothing Then Set vbComp = ActiveComp
    Dim n As Long
    Dim s As String
   With vbComp.CodeModule
        For n = .CountOfLines To 1 Step -1
            If .CountOfLines = 0 Then Exit For
            s = .Lines(n, 1)
            If Left(Trim(s), 5) = "Debug" Then
                LineString = s
                .REPLACELINE n, "'" & s
            ElseIf Left(Trim(s), 1) = "'" Then
            Else
            End If
        Next n
    End With
End Sub
Public Sub DebugEnableModule(Optional vbComp As VBComponent)
'#IMPORTS ActiveComp
If vbComp Is Nothing Then Set vbComp = ActiveComp
    Dim n As Long
    Dim s As String
   With vbComp.CodeModule
        For n = .CountOfLines To 1 Step -1
            If .CountOfLines = 0 Then Exit For
            s = .Lines(n, 1)
            If Left(Trim(s), 6) = "'Debug" Then
                LineString = s
                .REPLACELINE n, Mid(LineString, 2)
            ElseIf Left(Trim(s), 1) = "'" Then
            Else
            End If
        Next n
    End With
End Sub
Public Sub DebugEnableProcedure(Optional vbComp As VBComponent, Optional procName As String)
'#IMPORTS ActiveComp
'#IMPORTS ProcedureEndLine
'#IMPORTS ProcedureStartLine
'#IMPORTS ActiveProcName
If vbComp Is Nothing Then Set vbComp = ActiveComp
If procName = "" Then procName = ActiveProcName
    Dim n As Long
    Dim s As String
    With vbComp.CodeModule
        For n = ProcedureEndLine(vbComp.CodeModule, ActiveProcName) To _
                                                                    ProcedureStartLine(vbComp.CodeModule, procName) Step -1
            s = .Lines(n, 1)
            If Left(Trim(s), 6) = "'Debug" Then
                LineString = s
                .REPLACELINE n, Mid(LineString, 2)
            ElseIf Left(Trim(s), 1) = "'" Then
            Else
            End If
        Next n
    End With
End Sub

Public Sub DebugDisableProcedure(Optional vbComp As VBComponent, Optional procName As String)
'#IMPORTS ActiveComp
'#IMPORTS ProcedureEndLine
'#IMPORTS ProcedureStartLine
'#IMPORTS ActiveProcName
If vbComp Is Nothing Then Set vbComp = ActiveComp
If procName = "" Then procName = ActiveProcName
    Dim n As Long
    Dim s As String
    With vbComp.CodeModule
        For n = ProcedureEndLine(vbComp.CodeModule, ActiveProcName) To _
                                                                    ProcedureStartLine(vbComp.CodeModule, ActiveProcName) Step -1
            s = .Lines(n, 1)
            If Left(Trim(s), 5) = "Debug" Then
                LineString = s
                .REPLACELINE n, "'" & LineString
                Debug.Print LineString
            ElseIf Left(Trim(s), 1) = "'" Then
            Else
                LineCount = LineCount + 1
            End If
        Next n
    End With
End Sub

 '''
 Public Sub EncapsulateMultiple(Optional leftCapsule As String = """", Optional rightCapsule As String = """", Optional splitter As String = vbNewLine)
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    Dim arr
    arr = Split(code, splitter)
    Dim counter As Long
    For counter = LBound(arr) To UBound(arr) - IIf(Right(UBound(arr), Len(splitter)) = splitter, Len(splitter), 0)
        arr(counter) = leftCapsule & arr(counter) & rightCapsule
    Next
    code = Join(arr, splitter)
    
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & code & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub
'''
Public Sub FlipMultiple(Optional flipper As String = "=", Optional splitter As String = vbNewLine)
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    Dim arr
    arr = Split(code, splitter)
    Dim counter As Long
    For counter = LBound(arr) To UBound(arr) - IIf(Right(UBound(arr), Len(splitter)) = splitter, Len(splitter), 0)
        arr(counter) = Split(arr(counter), flipper)(1) & flipper & Split(arr(counter), flipper)(0)
    Next
    code = Join(arr, splitter)
    
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & code & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub
Public Sub Flip(Optional delim As String = "=")
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & Split(code, delim)(1) & delim & Split(code, delim)(0) & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub
'''

Public Sub sortSelection(delimeter As String)
'#IMPORTS sortSelectionArray
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    Dim arr
    arr = Split(code, delimeter)
    sortSelectionArray arr
    code = Join(arr, delimeter)
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & code & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub

Public Function sortSelectionArray(ByRef TempArray As Variant)
    Dim MaxVal As Variant
    Dim MaxIndex As Integer
    Dim I As Integer, J As Integer
    ' Step through the elements in the array starting with the
    ' last element in the array.
    For I = UBound(TempArray) To 0 Step -1

        ' Set MaxVal to the element in the array and save the
        ' index of this element as MaxIndex.
        MaxVal = TempArray(I)
        MaxIndex = I

        ' Loop through the remaining elements to see if any is
        ' larger than MaxVal. If it is then set this element
        ' to be the new MaxVal.
        For J = 0 To I
            If TempArray(J) > MaxVal Then
                MaxVal = TempArray(J)
                MaxIndex = J
            End If
        Next J

        ' If the index of the largest element is not i, then
        ' exchange this element with element i.
        If MaxIndex < I Then
            TempArray(MaxIndex) = TempArray(I)
            TempArray(I) = MaxVal
        End If
    Next I

End Function
'''

Public Sub CaseUpper()
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & UCase(code) & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub
Public Sub CaseLower()
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & LCase(code) & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub
Public Sub CaseProper()
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & WorksheetFunction.Proper(code) & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub

Public Sub Encapsulate(Optional before As String = """", Optional after As String = """")
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & before & code & after & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub


'''
Public Sub Inject(str As String)
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection   Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
   Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & str & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub
'''
Public Sub SubstituteInSelection(oldValue As String, newValue As String)
'#IMPORTS PartBeforeCodePaneSelection
'#IMPORTS PartAfterCodePaneSelection
'#IMPORTS CodepaneSelection
   Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    Dim code As String
    code = CodepaneSelection
    code = PartBeforeCodePaneSelection(startLine, StartColumn, EndLine, EndColumn) _
                & Replace(code, oldpart, newpart, , , vbTextCompare) & _
                PartAfterCodePaneSelection(startLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines startLine, EndLine - startLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines startLine, code
End Sub
'''

Public Function PartBeforeCodePaneSelection(startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long)
    Dim str As String
    str = Application.VBE.ActiveCodePane.CodeModule.Lines(startLine, 1)
    str = Mid(str, 1, StartColumn - 1)
    PartBeforeCodePaneSelection = str
End Function

Public Function PartAfterCodePaneSelection(startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long)
    Dim str As String
    str = Application.VBE.ActiveCodePane.CodeModule.Lines(EndLine, 1)
    str = Mid(str, EndColumn)
    PartAfterCodePaneSelection = str
End Function

Public Function CodepaneSelection() As String
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    If EndLine - startLine = 0 Then
        CodepaneSelection = Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(startLine, 1), StartColumn, EndColumn - StartColumn)
        Exit Function
    End If
    Dim str As String
    Dim I As Long
    For I = startLine To EndLine
        If str = "" Then
            str = Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(I, 1), StartColumn)
        ElseIf I < EndLine Then
            str = str & vbNewLine & Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(I, 1), 1)
        Else
            str = str & vbNewLine & Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(I, 1), 1, EndColumn - 1)
        End If
    Next
    CodepaneSelection = str
End Function

'''

Public Function vbeCursorPosition() As Long
    Dim startLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, EndLine, EndColumn
    vbeCursorPosition = startLine
End Function


'''

Public Sub ShowProcedureInfo(vbComp As VBComponent, procName As String)
'#IMPORTS ProcedureInfo
    Dim vbProj As VBIDE.VBProject
    Dim codeMod As VBIDE.CodeModule
    Dim CompName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim PInfo As ProcInfo
    CompName = vbComp.Name
    ProcKind = vbext_pk_Proc
    Set vbProj = ActiveWorkbook.VBProject
    Set vbComp = vbProj.VBComponents(CompName)
    Set codeMod = vbComp.CodeModule
    PInfo = ProcedureInfo(vbComp, procName, ProcKind)
    Debug.Print "ProcName: " & PInfo.procName
    Debug.Print "ProcKind: " & CStr(PInfo.ProcKind)
    Debug.Print "ProcStartLine: " & CStr(PInfo.ProcStartLine)
    Debug.Print "ProcBodyLine: " & CStr(PInfo.ProcBodyLine)
    Debug.Print "ProcCountLines: " & CStr(PInfo.ProcCountLines)
    Debug.Print "ProcScope: " & CStr(PInfo.ProcScope)
    Debug.Print "ProcDeclaration: " & PInfo.ProcDeclaration
End Sub

Public Function ProcedureInfo(vbComp As VBComponent, procName As String, ProcKind As VBIDE.vbext_ProcKind) As ProcInfo
'#IMPORTS GetProcedureDeclaration
    Dim codeMod As VBIDE.CodeModule
    Set codeMod = vbComp.CodeModule
    Dim PInfo As ProcInfo
    Dim BodyLine As Long
    Dim Declaration As String
    Dim FIRSTLINE As String
    BodyLine = codeMod.ProcStartLine(procName, ProcKind)
    If BodyLine > 0 Then
        With codeMod
            PInfo.procName = procName
            PInfo.ProcKind = ProcKind
            PInfo.ProcBodyLine = .ProcBodyLine(procName, ProcKind)
            PInfo.ProcCountLines = .ProcCountLines(procName, ProcKind)
            PInfo.ProcStartLine = .ProcStartLine(procName, ProcKind)
            FIRSTLINE = .Lines(PInfo.ProcBodyLine, 1)
            If StrComp(Left(FIRSTLINE, Len("Public")), "Public", vbBinaryCompare) = 0 Then
                PInfo.ProcScope = ScopePublic
            ElseIf StrComp(Left(FIRSTLINE, Len("public")), "public", vbBinaryCompare) = 0 Then
                PInfo.ProcScope = ScopePublic
            ElseIf StrComp(Left(FIRSTLINE, Len("Friend")), "Friend", vbBinaryCompare) = 0 Then
                PInfo.ProcScope = ScopeFriend
            Else
                PInfo.ProcScope = ScopeDefault
            End If
            PInfo.ProcDeclaration = GetProcedureDeclaration(vbComp, procName, ProcKind, LineSplitKeep)
        End With
    End If
    ProcedureInfo = PInfo
End Function

Public Function GetProcedureDeclaration(vbComp As VBComponent, _
                                        procName As String, ProcKind As VBIDE.vbext_ProcKind, _
                                        Optional LineSplitBehavior As LineSplits = LineSplitRemove)
'#IMPORTS SingleSpace
Dim codeMod As VBIDE.CodeModule
Set codeMod = vbComp.CodeModule
    Dim LineNum As Long
    Dim s As String
    Dim Declaration As String
    On Error Resume Next
    LineNum = codeMod.ProcBodyLine(procName, ProcKind)
    If Err.Number <> 0 Then
        Exit Function
    End If
    s = codeMod.Lines(LineNum, 1)
    Do While Right(s, 1) = "_"
        Select Case True
        Case LineSplitBehavior = LineSplitConvert
            s = Left(s, Len(s) - 1) & vbNewLine
        Case LineSplitBehavior = LineSplitKeep
            s = s & vbNewLine
        Case LineSplitBehavior = LineSplitRemove
            s = Left(s, Len(s) - 1) & " "
        End Select
        Declaration = Declaration & s
        LineNum = LineNum + 1
        s = codeMod.Lines(LineNum, 1)
    Loop
    Declaration = SingleSpace(Declaration & s)
    GetProcedureDeclaration = Declaration
End Function

Public Function SingleSpace(ByVal Text As String) As String
    Dim Pos As String
    Pos = InStr(1, Text, Space(2), vbBinaryCompare)
    Do Until Pos = 0
        Text = Replace(Text, Space(2), Space(1))
        Pos = InStr(1, Text, Space(2), vbBinaryCompare)
    Loop
    SingleSpace = Text
End Function





