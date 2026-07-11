Attribute VB_Name = "Module16"
Sub ToggleWeekTotal()
    Dim body As Shape, knob As Shape
    Dim pt As PivotTable
    Set body = ActiveSheet.Shapes("swBody2")
    Set knob = ActiveSheet.Shapes("swKnob2")
    Set pt = ActiveSheet.PivotTables(1)

    Dim leftPos As Double, rightPos As Double
    leftPos = body.Left + 4
    rightPos = body.Left + body.Width - knob.Width - 4

    ' Mevcut deger alanini temizle
    Dim df As PivotField
    For Each df In pt.DataFields
        df.Orientation = xlHidden
    Next df

    If knob.Left < (leftPos + rightPos) / 2 Then
        ' TOTAL'a geç
        knob.Left = rightPos
        knob.Fill.ForeColor.RGB = RGB(91, 155, 213)
        pt.AddDataField pt.PivotFields("TOTAL COST"), "Sum of TOTAL COST", xlSum
        ' yazi vurgusu
        ActiveSheet.Shapes("txtTotal").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ActiveSheet.Shapes("txtWeekly").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
    Else
        ' WEEKLY'ye geç
        knob.Left = leftPos
        knob.Fill.ForeColor.RGB = RGB(255, 192, 0)
        pt.AddDataField pt.PivotFields("WEEK COST"), "Sum of WEEK COST", xlSum
        ' yazi vurgusu
        ActiveSheet.Shapes("txtWeekly").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ActiveSheet.Shapes("txtTotal").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
    End If
End Sub
