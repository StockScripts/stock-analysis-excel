Attribute VB_Name = "Item3_Resources"
Option Explicit

Private Const QuickRatioIdeal = 2
Private Const QuickRatioMin = 1
Private Const QuickRatioMaxDecrease = -0.4

Public QuickRatio1 As Double
Public QuickRatio2 As Double
Public QuickRatio3 As Double
Public QuickRatio4 As Double
Public QuickRatio5 As Double

Sub QuickRatio()

    Dim ErrorNum As Integer
    
    On Error GoTo ErrorHandler

    Range("A12").Font.Bold = True
    Range("A12") = "Can they pay their bills?"
    
'   name Revenue cell
    Range("B13").Name = "QuickRatio"
    Range("13:13").Name = "QuickRatioRow"
    
'   write "Revenue" text
    Range("QuickRatio").HorizontalAlignment = xlLeft
    Range("QuickRatio") = "Quick Ratio"
    
    Range("QuickRatioRow").NumberFormat = "0.00"
    
    Range("QuickRatio").AddComment
    Range("QuickRatio").Comment.Visible = False
    Range("QuickRatio").Comment.Text Text:="quick ratio = (current assets - inventory) / current liabilities" & Chr(10) & _
                "must be > 2 and not decreasing" & Chr(10) & _
                "better measure than current ratio which includes inventory and is thus higher"
    Range("QuickRatio").Comment.Shape.TextFrame.AutoSize = True
    
'   populate QuickRatio information
    ErrorNum = 1
    QuickRatio1 = (CurrentAssets1 - Inventory1) / CurrentLiabilities1
    If QuickRatio1 >= QuickRatioIdeal Then
        Range("QuickRatio").Offset(0, 1).Font.ColorIndex = GreenFont
    ElseIf QuickRatio1 >= QuickRatioMin Then
        Range("QuickRatio").Offset(0, 1).Font.ColorIndex = OrangeFont
    Else
        Range("QuickRatio").Offset(0, 1).Font.ColorIndex = RedFont
    End If
    Range("QuickRatio").Offset(0, 1) = QuickRatio1
    
    ErrorNum = 2
    QuickRatio2 = (CurrentAssets1 - Inventory1) / CurrentLiabilities2
    If QuickRatio2 >= QuickRatioIdeal Then
        Range("QuickRatio").Offset(0, 2).Font.ColorIndex = GreenFont
    ElseIf QuickRatio2 >= QuickRatioMin Then
        Range("QuickRatio").Offset(0, 2).Font.ColorIndex = OrangeFont
    Else
        Range("QuickRatio").Offset(0, 2).Font.ColorIndex = RedFont
    End If
    Range("QuickRatio").Offset(0, 2) = QuickRatio2
    
    ErrorNum = 3
    QuickRatio3 = (CurrentAssets1 - Inventory1) / CurrentLiabilities3
    If QuickRatio3 >= QuickRatioIdeal Then
        Range("QuickRatio").Offset(0, 3).Font.ColorIndex = GreenFont
    ElseIf QuickRatio3 >= QuickRatioMin Then
        Range("QuickRatio").Offset(0, 3).Font.ColorIndex = OrangeFont
    Else
        Range("QuickRatio").Offset(0, 3).Font.ColorIndex = RedFont
    End If
    Range("QuickRatio").Offset(0, 3) = QuickRatio3
    
    ErrorNum = 4
    QuickRatio4 = (CurrentAssets1 - Inventory1) / CurrentLiabilities4
    If QuickRatio4 >= QuickRatioIdeal Then
        Range("QuickRatio").Offset(0, 4).Font.ColorIndex = GreenFont
    ElseIf QuickRatio4 >= QuickRatioMin Then
        Range("QuickRatio").Offset(0, 4).Font.ColorIndex = OrangeFont
    Else
        Range("QuickRatio").Offset(0, 4).Font.ColorIndex = RedFont
    End If
    Range("QuickRatio").Offset(0, 4) = QuickRatio4
    
    ErrorNum = 5
    QuickRatio5 = (CurrentAssets1 - Inventory1) / CurrentLiabilities5
    If QuickRatio5 >= QuickRatioIdeal Then
        Range("QuickRatio").Offset(0, 5).Font.ColorIndex = GreenFont
    ElseIf QuickRatio5 >= QuickRatioMin Then
        Range("QuickRatio").Offset(0, 5).Font.ColorIndex = OrangeFont
    Else
        Range("QuickRatio").Offset(0, 5).Font.ColorIndex = RedFont
    End If
    Range("QuickRatio").Offset(0, 5) = QuickRatio5
    
'   calculate YOY Working Capital growth
    QuickRatioYOY
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case 1
            QuickRatio1 = 0
            Range("QuickRatio").Offset(0, 1) = QuickRatio1
        Case 2
            QuickRatio2 = 0
            Range("QuickRatio").Offset(0, 2) = QuickRatio2
        Case 3
            QuickRatio3 = 0
            Range("QuickRatio").Offset(0, 3) = QuickRatio3
        Case 4
            QuickRatio4 = 0
            Range("QuickRatio").Offset(0, 4) = QuickRatio4
        Case 5
            QuickRatio5 = 0
            Range("QuickRatio").Offset(0, 5) = QuickRatio5
   End Select
    
   QuickRatioYOY

End Sub

Sub QuickRatioYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B14").Name = "YOYGrowth"
    Range("14:14").Name = "YOYRow"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("14:14").NumberFormat = "0.0%"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(QuickRatio1, QuickRatio2)
    YOY2 = YOYGrowth(QuickRatio2, QuickRatio3)
    YOY3 = YOYGrowth(QuickRatio3, QuickRatio4)
    YOY4 = YOYGrowth(QuickRatio4, QuickRatio5)
    
    Call YOYQuickRatioEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYQuickRatioEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    If QuickRatio4 < 0 Or YOY4 < QuickRatioMaxDecrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY4 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    If QuickRatio3 < 0 Or YOY3 < QuickRatioMaxDecrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY3 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If QuickRatio2 < 0 Or YOY2 < QuickRatioMaxDecrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY2 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If QuickRatio1 < 0 Or YOY1 < QuickRatioMaxDecrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY1 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function


