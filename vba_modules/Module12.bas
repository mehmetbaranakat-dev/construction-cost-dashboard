Attribute VB_Name = "Module12"
' ===== FILTRE BUTONLARI =====

Sub FilterDiesel()
    Call ClearFiltersMainSheet
    FilterCostItem "Diesel"
End Sub

Sub FilterDrinkingWater()
    Call ClearFiltersMainSheet
    FilterCostItem "Drinking Water"
End Sub

Sub FilterEffluent()
    Call ClearFiltersMainSheet
    FilterCostItem "Effluent collection"
End Sub

' ===== ORTAK YARDIMCI =====

Sub FilterCostItem(criteria As String)
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects(1)
    
    tbl.Range.AutoFilter Field:=tbl.ListColumns("ADDRESSING THE COST ITEM").Index, _
                         Criteria1:=criteria
End Sub

