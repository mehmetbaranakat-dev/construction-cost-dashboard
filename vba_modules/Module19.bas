Attribute VB_Name = "Module21"
Sub CompanySelected()
    Dim idx As Long
    idx = ThisWorkbook.Sheets("CHART1- MATERIAL&SERVICE BASED").Range("AZ1").Value
    If idx >= 1 Then
        ThisWorkbook.Sheets("CHART1- MATERIAL&SERVICE BASED").Range("B3").Value = _
            ThisWorkbook.Sheets("SUPPLIERS").Range("B1").Offset(idx, 0).Value
    End If
End Sub

