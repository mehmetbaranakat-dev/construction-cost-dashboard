Attribute VB_Name = "Module22"
Sub GoToChart1()
    ' Önce hedef sekmeyi göster ve aktif et
    ThisWorkbook.Sheets("CHART1- MATERIAL&SERVICE BASED").Visible = xlSheetVisible
    ThisWorkbook.Sheets("CHART1- MATERIAL&SERVICE BASED").Activate
    
    ' Sonra dashboard sekmesini gizle
    ThisWorkbook.Sheets("DASHBOARD").Visible = xlSheetHidden
End Sub
