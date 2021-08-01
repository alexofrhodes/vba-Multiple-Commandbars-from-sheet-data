Attribute VB_Name = "var"
Option Explicit

Sub ImageFromExternalFile()
    Dim Btn As Office.CommandBarButton
    Set Btn = Application.CommandBars.FindControl(ID:=30007) _
        .Controls.Add(Type:=msoControlButton, Temporary:=True)
    With Btn
        .Caption = "Click Me"
        .Style = msoButtonIconAndCaption
        .Picture = LoadPicture("C:\TestPic.bmp")
    End With
End Sub

Sub ImageFromEmbedded()
    Dim P As Excel.Picture
    Dim Btn As Office.CommandBarButton
    Set Btn = Application.CommandBars.FindControl(ID:=30007) _
        .Controls.Add(Type:=msoControlButton, Temporary:=True)
    Set P = Worksheets("Sheet1").Pictures("ThePict")
    P.CopyPicture xlScreen, xlBitmap
    With Btn
        .Caption = "Click Me"
        .Style = msoButtonIconAndCaption
        .PasteFace
    End With
End Sub

Sub TestCallingSubFromFile()
    Application.Run "'" & Workbooks("test.xlsm").FullName & "'!Testme"
End Sub

Sub ResetCBAR()
'eg
    Excel.Application.CommandBars("Cell").Reset
End Sub


