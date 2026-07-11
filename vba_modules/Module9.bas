Attribute VB_Name = "Module7"
Sub ClearFiltersMainSheet()
    Dim tbl As ListObject
    Dim lastTbl As ListObject
    Dim keyCol As Range
    Dim lastDataRow As Long
    
    ' Aktif sayfadaki tüm tablolarda filtreleri kaldir
    For Each tbl In ActiveSheet.ListObjects
        If tbl.ShowAutoFilter Then
            If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
        End If
        Set lastTbl = tbl
    Next tbl
    
    If lastTbl Is Nothing Then Exit Sub
    If lastTbl.DataBodyRange Is Nothing Then Exit Sub
    
    ' DATE sütununda dolu olan son satiri bul
    Set keyCol = lastTbl.ListColumns("DATE").DataBodyRange
    lastDataRow = keyCol.Cells.Find(What:="*", LookIn:=xlValues, _
                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    ' Bir altindaki satirin ilk hücresine git
    ActiveSheet.Cells(lastDataRow + 1, lastTbl.Range.Columns(1).Column).Select
End Sub

