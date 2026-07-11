Attribute VB_Name = "Module3"
Sub GoToDashboard()
    Dim suanki As Worksheet
    Set suanki = ActiveSheet
    
    Sheets("DASHBOARD").Visible = True
    Sheets("DASHBOARD").Activate
    Range("A1").Select
    
    If suanki.Name <> "DASHBOARD" Then suanki.Visible = xlSheetHidden
End Sub
