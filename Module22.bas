Attribute VB_Name = "Module10"
Sub RefreshAllQueries()
    Dim lConn As WorkbookConnection
    
    For Each lConn In ThisWorkbook.Connections
        If Not lConn.OLEDBConnection Is Nothing Then
            lConn.OLEDBConnection.BackgroundQuery = False
        End If
    Next lConn
    
    ThisWorkbook.RefreshAll
    Application.CAlculate
    
    ' Sifir HGV etiketlerini gizle
    Dim cht As Chart, s As Series, i As Long
    Dim v As Variant
    Set cht = Sheets("HGV").ChartObjects("Chart 9").Chart
    For Each s In cht.SeriesCollection
        v = s.Values
        For i = 1 To UBound(v)
            s.Points(i).HasDataLabel = True
            If v(i) = 0 Then s.Points(i).HasDataLabel = False
        Next i
    Next s
    
    MsgBox "All queries have been refreshed successfully.", vbInformation, "Update Complete"
End Sub
