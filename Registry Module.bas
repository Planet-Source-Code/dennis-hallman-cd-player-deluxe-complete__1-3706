Attribute VB_Name = "RegistrySet"

Public Sub SaveWindowPos(ByVal frm As Form)
    ' Save window position and size attributes to registry
    SaveSetting App.Title, "Config", frm.Name + "_left", CStr(frm.Left)
    SaveSetting App.Title, "Config", frm.Name + "_top", CStr(frm.Top)
    SaveSetting App.Title, "Config", frm.Name + "_width", CStr(frm.Width)
    SaveSetting App.Title, "Config", frm.Name + "_height", CStr(frm.Height)
End Sub

Public Sub LoadWindowPos(ByVal frm As Form)
    ' See if any settings are saved for this window
    If GetSetting(App.Title, "Config", frm.Name + "_left", "") = "" Then
        ' There aren't - so just centre the form on the screen
        frm.Left = (Screen.Width - frm.Width) / 2
        frm.Top = (Screen.Height - frm.Height) / 2
        Exit Sub
    End If
    ' Load the form's attributes from the registry
    frm.Left = CLng(GetSetting(App.Title, "Config", frm.Name + "_left", "0"))
    frm.Top = CLng(GetSetting(App.Title, "Config", frm.Name + "_top", "0"))
    frm.Width = CLng(GetSetting(App.Title, "Config", frm.Name + "_width", CStr(frm.Width)))
    frm.Height = CLng(GetSetting(App.Title, "Config", frm.Name + "_height", CStr(frm.Height)))
    ' Optional - if the form is opening with any part off the screen
    '     then nudge it back on
    If frm.Left < 0 Then frm.Left = 0
    If frm.Top < 0 Then frm.Top = 0
    If frm.Left + frm.Width > Screen.Width Then frm.Left = Screen.Width - frm.Width
    If frm.Top + frm.Height > Screen.Height Then frm.Top = Screen.Height - frm.Height
End Sub

Public Sub SaveColSet(ByVal frm As Form, dcol As Long)
    'Save Program Color Setting. (This Works).
    'SaveSetting App.Title, "BkgColor", "Color", DenColor
    SaveSetting appname:=App.Title, section:="BackColor", Key:=frm.Name + "_Color", setting:=dcol
End Sub

Public Sub LoadColSet(ByVal frm As Form)
    If GetSetting(App.Title, "Backcolor", frm.Name + "_Color", "") = "" Then
        'Default Program Color Setting. (This Works)
        frm.TimeWindow.ForeColor = &HFF8080
        frm.TotalPlay.ForeColor = &HFF8080
        frm.TrackTime.ForeColor = &HFF8080
        frm.cboTrack.ForeColor = &HFF8080
        frm.Label1.ForeColor = &HFF8080
        Exit Sub
    End If
    'Get Program Color Setting. (This Works)
    frm.TimeWindow.ForeColor = GetSetting(App.Title, "BackColor", frm.Name + "_Color")
    frm.TotalPlay.ForeColor = GetSetting(App.Title, "BackColor", frm.Name + "_Color")
    frm.TrackTime.ForeColor = GetSetting(App.Title, "BackColor", frm.Name + "_Color")
    frm.cboTrack.ForeColor = GetSetting(App.Title, "BackColor", frm.Name + "_Color")
    frm.Label1.ForeColor = GetSetting(App.Title, "BackColor", frm.Name + "_Color")
End Sub

Public Sub SaveSkinSet(ByVal frm As Form, Dskin As String)
    'Save Program Skin Setting. (This Works).
    SaveSetting appname:=App.Title, section:="Skin", Key:=frm.Name + "_Skin", setting:=Trim(Dskin)
End Sub

Public Sub LoadSkinSet(ByVal frm As Form)
    If GetSetting(App.Title, "Skin", frm.Name + "_Skin", "") = "" Then
        Exit Sub
    End If
    frm.PicSourceImage.Picture = LoadPicture(Trim(GetSetting(App.Title, "Skin", frm.Name + "_Skin")))
End Sub

