VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Private Sub Workbook_Open()
runListProcs
End Sub



'----NOTES----
'-------------
'Private Sub Workbook_Activate()
'
''call different menu based on current user
'    Dim sUserName As String
'    sUserName = Application.UserName
'    Select Case sUserName
'    Case "Ron de Bruin": Call AddToCellMenu
'    'Case "Dave Thomson": Call AddToCellMenu2
'    'Case Else: Call AddToCellMenu3
'    End Select


''call different menu based on current language
'    Dim LangID As Long
'    LangID = Application.International(xlCountryCode)
'    Select Case LangID
'    Case 31: 'Call AddToCellMenuInDutch
'    Case 49: 'Call AddToCellMenuinGerman
'    Case Else: Call AddToCellMenu
'    End Select


'End Sub
