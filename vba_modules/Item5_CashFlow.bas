Attribute VB_Name = "Item5_CashFlow"
Private Const CashFlowMaxDecrease = -0.2

Sub FreeCashFlow()

    Range("A20").Font.Bold = True
    Range("A20") = "Do they have good cash flow?"
    
'   name OpCashFlow cell
    Range("B21").Name = "FreeCashFlow"
    
'   write "Operating Cash Flow" text
    Range("FreeCashFlow").HorizontalAlignment = xlLeft
    Range("FreeCashFlow") = "Free Cash Flow"
    
'   populate revenue information
    If FreeCashFlow1 >= 0 Then
        Range("FreeCashFlow").Offset(0, 1).Font.ColorIndex = GreenFont
    Else
        Range("FreeCashFlow").Offset(0, 1).Font.ColorIndex = RedFont
    End If
    Range("FreeCashFlow").Offset(0, 1) = FreeCashFlow1
    
    If FreeCashFlow2 >= 0 Then
        Range("FreeCashFlow").Offset(0, 2).Font.ColorIndex = GreenFont
    Else
        Range("FreeCashFlow").Offset(0, 2).Font.ColorIndex = RedFont
    End If
    Range("FreeCashFlow").Offset(0, 2) = FreeCashFlow2
    
    If FreeCashFlow3 >= 0 Then
        Range("FreeCashFlow").Offset(0, 3).Font.ColorIndex = GreenFont
    Else
        Range("FreeCashFlow").Offset(0, 3).Font.ColorIndex = RedFont
    End If
    Range("FreeCashFlow").Offset(0, 3) = FreeCashFlow3
    
    If FreeCashFlow4 >= 0 Then
        Range("FreeCashFlow").Offset(0, 4).Font.ColorIndex = GreenFont
    Else
        Range("FreeCashFlow").Offset(0, 4).Font.ColorIndex = RedFont
    End If
    Range("FreeCashFlow").Offset(0, 4) = FreeCashFlow4
    
    If FreeCashFlow5 >= 0 Then
        Range("FreeCashFlow").Offset(0, 5).Font.ColorIndex = GreenFont
    Else
        Range("FreeCashFlow").Offset(0, 5).Font.ColorIndex = RedFont
    End If
    Range("FreeCashFlow").Offset(0, 5) = FreeCashFlow5
    
    Range("FreeCashFlow").AddComment
    Range("FreeCashFlow").Comment.Visible = False
    Range("FreeCashFlow").Comment.Text Text:="operating cash flow - capital expenses" & Chr(10) & _
                "should be positive or increasing"
    Range("FreeCashFlow").Comment.Shape.TextFrame.AutoSize = True
    
'   calculate YOY revenue growth
    FreeCashFlowYOY

End Sub

Sub FreeCashFlowYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B22").Name = "YOYGrowth"
    Range("22:22").Name = "YOYRow"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").NumberFormat = "0.0%"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(FreeCashFlow1, FreeCashFlow2)
    YOY2 = YOYGrowth(FreeCashFlow2, FreeCashFlow3)
    YOY3 = YOYGrowth(FreeCashFlow3, FreeCashFlow4)
    YOY4 = YOYGrowth(FreeCashFlow4, FreeCashFlow5)
    
    Call YOYFreeCashFlowEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYFreeCashFlowEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    If FreeCashFlow4 < 0 Or YOY4 < CashFlowMaxDecrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY4 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    If FreeCashFlow3 < 0 Or YOY3 < CashFlowMaxDecrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY3 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If FreeCashFlow2 < 0 Or YOY2 < CashFlowMaxDecrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY2 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If FreeCashFlow1 < 0 Or YOY1 < CashFlowMaxDecrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY1 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function

