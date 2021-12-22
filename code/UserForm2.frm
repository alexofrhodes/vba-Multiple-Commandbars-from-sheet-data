VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   5508
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2076
   OleObjectBlob   =   "UserForm2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub TextBox1_Change()

End Sub

Public Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 2 Then
        If BarExists("rClickBar") Then
            CommandBars("rClickBar").ShowPopup
            Cancel = True
        End If
    End If
End Sub

Public Sub UserForm_Click()

End Sub

