Attribute VB_Name = "Module1"
Sub exampleOfControls()
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
    
    ''the following doesn't work
'    Set cbc = .Controls.Add(Type:=msoControlEdit, Temporary:=True)
'    cbc.Caption = "Text box:"
'    cbc.Text = "Type in a text:"
'    cbc.OnAction = "MenuText_OnAction"
    
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
