Attribute VB_Name = "Module6"
Sub GoToHGV()
    Dim suanki As Worksheet
    Set suanki = ActiveSheet
    
    Sheets("HGV").Visible = True
    Sheets("HGV").Activate
    Range("A1").Select
    
    If suanki.Name <> "HGV" Then suanki.Visible = xlSheetHidden
End Sub
