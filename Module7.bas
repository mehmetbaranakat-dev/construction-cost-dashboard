Attribute VB_Name = "Module5"
Sub GoToPrincipleBased()
    Dim suanki As Worksheet
    Set suanki = ActiveSheet
    
    Sheets("PRINCIPLE BASED").Visible = True
    Sheets("PRINCIPLE BASED").Activate
    Range("A1").Select
    
    If suanki.Name <> "PRINCIPLE BASED" Then suanki.Visible = xlSheetHidden
End Sub
