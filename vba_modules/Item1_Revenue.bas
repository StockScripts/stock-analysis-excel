Attribute VB_Name = "Item1_Revenue"
Option Explicit

Private Const SlowRevenueGrowth = 0.1

Sub Revenue()

    Range("A2").Font.Bold = True
    Range("A2") = "Are they making sales?"
    
'   name Revenue cell
    Range("B3").Name = "Revenue"
    
'   write "Revenue" text
    Range("Revenue").HorizontalAlignment = xlLeft
    Range("Revenue") = "Revenue"
    
'   populate revenue information
    If Revenue1 >= 0 Then
        Range("Revenue").Offset(0, 1).Font.ColorIndex = GreenFont
    Else
        Range("Revenue").Offset(0, 1).Font.ColorIndex = RedFont
    End If
    Range("Revenue").Offset(0, 1) = Revenue1
    
    If Revenue2 >= 0 Then
        Range("Revenue").Offset(0, 2).Font.ColorIndex = GreenFont
    Else
        Range("Revenue").Offset(0, 2).Font.ColorIndex = RedFont
    End If
    Range("Revenue").Offset(0, 2) = Revenue2
    
    If Revenue3 >= 0 Then
        Range("Revenue").Offset(0, 3).Font.ColorIndex = GreenFont
    Else
        Range("Revenue").Offset(0, 3).Font.ColorIndex = RedFont
    End If
    Range("Revenue").Offset(0, 3) = Revenue3
    
    If Revenue4 >= 0 Then
        Range("Revenue").Offset(0, 4).Font.ColorIndex = GreenFont
    Else
        Range("Revenue").Offset(0, 4).Font.ColorIndex = RedFont
    End If
    Range("Revenue").Offset(0, 4) = Revenue4
    
    If Revenue5 >= 0 Then
        Range("Revenue").Offset(0, 5).Font.ColorIndex = GreenFont
    Else
        Range("Revenue").Offset(0, 5).Font.ColorIndex = RedFont
    End If
    Range("Revenue").Offset(0, 5) = Revenue5
    
    Range("Revenue").AddComment
    Range("Revenue").Comment.Visible = False
    Range("Revenue").Comment.Text Text:="must rise with net profit margin" & Chr(10) & _
                "to increase earnings"
    Range("Revenue").Comment.Shape.TextFrame.AutoSize = True
    
'   calculate YOY revenue growth
    RevenueYOY

End Sub

Sub RevenueYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B4").Name = "YOYGrowth"
    Range("4:4").Name = "YOYRow"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    Range("YOYRow").NumberFormat = "0.0%"
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(Revenue1, Revenue2)
    YOY2 = YOYGrowth(Revenue2, Revenue3)
    YOY3 = YOYGrowth(Revenue3, Revenue4)
    YOY4 = YOYGrowth(Revenue4, Revenue5)
    
    Call YOYRevenueEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYRevenueEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    If Revenue4 < 0 Or YOY4 < 0 Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    If Revenue3 < 0 Or YOY3 < 0 Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY4 - YOY3) > SlowRevenueGrowth Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If Revenue2 < 0 Or YOY2 < 0 Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY3 - YOY2) > SlowRevenueGrowth Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If Revenue1 < 0 Or YOY1 < 0 Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY2 - YOY1) > SlowRevenueGrowth Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function

