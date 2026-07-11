Attribute VB_Name = "Sheet5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    
    For Each cell In Target
        If cell.Column = 13 Then ' M column
            If cell.Value <> "" And cell.Value <> 0 Then
                If Me.Cells(cell.Row, 8).Value = "Muck Away backloaded" Then
                    MsgBox "Warning: Backloaded trucks are already included in HGV count! You should not enter data in this field.", vbExclamation, "Duplicate Entry Warning"
                End If
            End If
        End If
    Next cell
End Sub
