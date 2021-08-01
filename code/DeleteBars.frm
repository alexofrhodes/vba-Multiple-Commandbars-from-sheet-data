VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteBars 
   Caption         =   "Delete custom bars"
   ClientHeight    =   4128
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2400
   OleObjectBlob   =   "DeleteBars.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If ListBox1.ListIndex = -1 Then Exit Sub
    C_TAG = ListBox1.List(ListBox1.ListIndex)
    DeleteAll
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("combarTAGS").Columns(1)
    Set rng = rng.Find(C_TAG)
    rng.Delete Shift:=xlUp
    ListBox1.RemoveItem ListBox1.ListIndex
End Sub

Private Sub UserForm_Initialize()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("combarTAGS").Range("A1").CurrentRegion
    If rng.Cells.Count = 1 Then Exit Sub
    For Each cell In rng.Cells
        If cell.Row <> 1 Then ListBox1.AddItem cell
    Next cell
End Sub

