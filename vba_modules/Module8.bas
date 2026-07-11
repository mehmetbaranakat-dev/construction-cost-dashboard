Attribute VB_Name = "Module4"
Sub GoToInventory()
    Sheets("INVENTORY LIST").Visible = True
    Sheets("INVENTORY LIST").Activate
    Range("A1").Select
End Sub
