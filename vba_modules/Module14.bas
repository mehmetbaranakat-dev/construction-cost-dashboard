Attribute VB_Name = "Module14"
Sub ToggleSwitch()
    Dim body As Shape, knob As Shape
    Set body = ActiveSheet.Shapes("swBody")
    Set knob = ActiveSheet.Shapes("swKnob")

    Dim leftPos As Double, rightPos As Double
    leftPos = body.Left + 4
    rightPos = body.Left + body.Width - knob.Width - 4

    ' Kapsül sabit beyaz (header'daki gibi)
    body.Fill.ForeColor.RGB = RGB(255, 255, 255)

    If knob.Left < (leftPos + rightPos) / 2 Then
        ' PV'ye geç -> topuz mavi
        knob.Left = rightPos
        knob.Fill.ForeColor.RGB = RGB(91, 155, 213)
        Sheets("PIE CHART BESS PV").Range("E23").Value = "PV"
    Else
        ' BESS'e geç -> topuz sari
        knob.Left = leftPos
        knob.Fill.ForeColor.RGB = RGB(255, 192, 0)
        Sheets("PIE CHART BESS PV").Range("E23").Value = "BESS"
    End If
    knob.TextFrame2.TextRange.Text = ChrW(8646)
    With knob.TextFrame2.TextRange.Font
        .Fill.ForeColor.RGB = RGB(255, 102, 102)
        .Size = 16
        .Bold = msoTrue
    End With
    knob.TextFrame2.VerticalAnchor = msoAnchorMiddle
    knob.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
  
End Sub
