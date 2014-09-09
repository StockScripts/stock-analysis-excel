Attribute VB_Name = "Item8_Dividend"

Sub DividendPerShare()

'   name DividendPerShare cell
    Range("B34").Name = "DividendPerShare"
    
'   write "Dividend/Share" text
    Range("DividendPerShare").HorizontalAlignment = xlLeft
    Range("DividendPerShare") = "Dividend/Share"
    
'   populate revenue information
    Range("DividendPerShare").Offset(0, 1) = DividendPerShare1
    Range("DividendPerShare").Offset(0, 2) = DividendPerShare2
    Range("DividendPerShare").Offset(0, 3) = DividendPerShare3
    Range("DividendPerShare").Offset(0, 4) = DividendPerShare4
    Range("DividendPerShare").Offset(0, 5) = DividendPerShare5
    
    Range("DividendPerShare").AddComment
    Range("DividendPerShare").Comment.Visible = False
    Range("DividendPerShare").Comment.Text Text:="earnings = sales x profit margin" & Chr(10) & _
                "to increase earnings"
    Range("DividendPerShare").Comment.Shape.TextFrame.AutoSize = True
    
'   calculate YOY Dividend Per Share growth
    DividendPerShareYOY
    
End Sub

Sub DividendPerShareYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B35").Name = "YOYGrowth"
    Range("35:35").Name = "YOYRow"
    
    Range("YOYRow").NumberFormat = "0.0%"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(DividendPerShare1, DividendPerShare2)
    YOY2 = YOYGrowth(DividendPerShare2, DividendPerShare3)
    YOY3 = YOYGrowth(DividendPerShare3, DividendPerShare4)
    YOY4 = YOYGrowth(DividendPerShare4, DividendPerShare5)
    
    Call YOYDivPerShareEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYDivPerShareEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    If YOY4 < 0 Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    If YOY3 < 0 Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If YOY2 < 0 Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If YOY1 < 0 Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function



