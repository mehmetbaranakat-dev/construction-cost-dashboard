Attribute VB_Name = "Module2"
Sub Update()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache
    
    ' Tüm Pivot Cache'leri güncelle
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc
    
    ' Tüm formülleri hesapla
    Application.CAlculate
    
    MsgBox "Updated!", vbInformation

End Sub
