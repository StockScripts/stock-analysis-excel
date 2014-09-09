Attribute VB_Name = "Item2_Profit"
Option Explicit

Private Const SlowNetIncomeGrowth = 0.5
Private Const NetIncomeMaxDecline = -0.2

Private Const SlowNetMarginGrowth = 0.5
Private Const NetMarginMaxDecline = -0.2

Private Const ROERequirement = 0.1
Private Const SlowROEGrowth = 0.2
Private Const ROEMaxDecline = -0.2

Public NetMargin1 As Double
Public NetMargin2 As Double
Public NetMargin3 As Double
Public NetMargin4 As Double
Public NetMargin5 As Double

Public ROE1 As Double
Public ROE2 As Double
Public ROE3 As Double
Public ROE4 As Double
Public ROE5 As Double

Sub Profit()

    NetIncome
    NetMargin
    ROE
    
End Sub

Sub NetIncome()

    Range("A5").Font.Bold = True
    Range("A5") = "Are they profitable?"
    
'   name Revenue cell
    Range("B6").Name = "NetIncome"
    
'   write "Revenue" text
    Range("NetIncome").HorizontalAlignment = xlLeft
    Range("NetIncome") = "Net Income"
    
'   populate revenue information
    If NetIncome1 >= 0 Then
        Range("NetIncome").Offset(0, 1).Font.ColorIndex = GreenFont
    Else
        Range("NetIncome").Offset(0, 1).Font.ColorIndex = RedFont
    End If
    Range("NetIncome").Offset(0, 1) = NetIncome1
    
    If NetIncome2 >= 0 Then
        Range("NetIncome").Offset(0, 2).Font.ColorIndex = GreenFont
    Else
        Range("NetIncome").Offset(0, 2).Font.ColorIndex = RedFont
    End If
    Range("NetIncome").Offset(0, 2) = NetIncome2
    
    If NetIncome3 >= 0 Then
        Range("NetIncome").Offset(0, 3).Font.ColorIndex = GreenFont
    Else
        Range("NetIncome").Offset(0, 3).Font.ColorIndex = RedFont
    End If
    Range("NetIncome").Offset(0, 3) = NetIncome3
    
    If NetIncome4 >= 0 Then
        Range("NetIncome").Offset(0, 4).Font.ColorIndex = GreenFont
    Else
        Range("NetIncome").Offset(0, 4).Font.ColorIndex = RedFont
    End If
    Range("NetIncome").Offset(0, 4) = NetIncome4
    
    If NetIncome5 >= 0 Then
        Range("NetIncome").Offset(0, 5).Font.ColorIndex = GreenFont
    Else
        Range("NetIncome").Offset(0, 5).Font.ColorIndex = RedFont
    End If
    Range("NetIncome").Offset(0, 5) = NetIncome5
    
    Range("NetIncome").AddComment
    Range("NetIncome").Comment.Visible = False
    Range("NetIncome").Comment.Text Text:="net income = operating income - interest expenses - income taxes" & Chr(10) & _
                "must increase faster than sales for earnings to increase"
    Range("NetIncome").Comment.Shape.TextFrame.AutoSize = True
    
'   calculate YOY net income growth
    NetIncomeYOY

End Sub

Sub NetIncomeYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B7").Name = "YOYGrowth"
    Range("7:7").Name = "YOYRow"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(NetIncome1, NetIncome2)
    YOY2 = YOYGrowth(NetIncome2, NetIncome3)
    YOY3 = YOYGrowth(NetIncome3, NetIncome4)
    YOY4 = YOYGrowth(NetIncome4, NetIncome5)
    
    Call YOYNetIncomeEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYNetIncomeEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    Selection.NumberFormat = "0.0%"
    If NetIncome4 < 0 Or YOY4 < NetIncomeMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    Selection.NumberFormat = "0.0%"
    If NetIncome3 < 0 Or YOY3 < NetIncomeMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY4 - YOY3) > SlowNetIncomeGrowth Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    Selection.NumberFormat = "0.0%"
    If NetIncome2 < 0 Or YOY2 < NetIncomeMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY3 - YOY2) > SlowNetIncomeGrowth Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    Selection.NumberFormat = "0.0%"
    If NetIncome1 < 0 Or YOY1 < NetIncomeMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY2 - YOY1) > SlowNetIncomeGrowth Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function

Sub NetMargin()

    Dim ErrorNum As Integer
    
    On Error GoTo ErrorHandler
    
'   name Net Margin cell
    Range("B8").Name = "NetMargin"
    Range("8:8").Name = "NetMarginRow"
    
'   write "Net Margin" text
    Range("NetMargin").HorizontalAlignment = xlLeft
    Range("NetMargin") = "Net Margin"
    
    Range("NetMarginRow").NumberFormat = "0.0%"
    
    Range("NetMargin").AddComment
    Range("NetMargin").Comment.Visible = False
    Range("NetMargin").Comment.Text Text:="net income/sales" & Chr(10) & _
                "must rise faster than revenue to increase earnings" & Chr(10) & _
                "must be increasing or at least stable"
    Range("NetMargin").Comment.Shape.TextFrame.AutoSize = True
    
    ErrorNum = 1
    NetMargin1 = NetIncome1 / Revenue1
    If NetMargin1 >= 0 Then
        Range("NetMargin").Offset(0, 1).Font.ColorIndex = GreenFont
    Else
        Range("NetMargin").Offset(0, 1).Font.ColorIndex = RedFont
    End If
    Range("NetMargin").Offset(0, 1) = NetMargin1
    
    ErrorNum = 2
    NetMargin2 = NetIncome2 / Revenue2
    If NetMargin2 >= 0 Then
        Range("NetMargin").Offset(0, 2).Font.ColorIndex = GreenFont
    Else
        Range("NetMargin").Offset(0, 2).Font.ColorIndex = RedFont
    End If
    Range("NetMargin").Offset(0, 2) = NetMargin2
    
    ErrorNum = 3
    NetMargin3 = NetIncome3 / Revenue3
    If NetMargin3 >= 0 Then
        Range("NetMargin").Offset(0, 3).Font.ColorIndex = GreenFont
    Else
        Range("NetMargin").Offset(0, 3).Font.ColorIndex = RedFont
    End If
    Range("NetMargin").Offset(0, 3) = NetMargin3
    
    ErrorNum = 4
    NetMargin4 = NetIncome4 / Revenue4
    If NetMargin4 >= 0 Then
        Range("NetMargin").Offset(0, 4).Font.ColorIndex = GreenFont
    Else
        Range("NetMargin").Offset(0, 4).Font.ColorIndex = RedFont
    End If
    Range("NetMargin").Offset(0, 4) = NetMargin4
    
    ErrorNum = 5
    NetMargin5 = NetIncome5 / Revenue5
    If NetMargin5 >= 0 Then
        Range("NetMargin").Offset(0, 5).Font.ColorIndex = GreenFont
    Else
        Range("NetMargin").Offset(0, 5).Font.ColorIndex = RedFont
    End If
    Range("NetMargin").Offset(0, 5) = NetMargin5
    
'   calculate YOY net margin growth
    NetMarginYOY
    
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case 1
            NetMargin1 = 0
            Range("NetMargin").Offset(0, 1) = NetMargin1
        Case 2
            NetMargin2 = 0
            Range("NetMargin").Offset(0, 2) = NetMargin2
        Case 3
            NetMargin3 = 0
            Range("NetMargin").Offset(0, 3) = NetMargin3
        Case 4
            NetMargin4 = 0
            Range("NetMargin").Offset(0, 4) = NetMargin4
        Case 5
            NetMargin5 = 0
            Range("NetMargin").Offset(0, 5) = NetMargin5
   End Select
   
   NetMarginYOY
    
End Sub

Sub NetMarginYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B9").Name = "YOYGrowth"
    Range("9:9").Name = "YOYRow"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With
    
    YOY1 = YOYGrowth(NetMargin1, NetMargin2)
    YOY2 = YOYGrowth(NetMargin2, NetMargin3)
    YOY3 = YOYGrowth(NetMargin3, NetMargin4)
    YOY4 = YOYGrowth(NetMargin4, NetMargin5)
    
    Call YOYNetMarginEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYNetMarginEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    Selection.NumberFormat = "0.0%"
    If NetMargin4 < 0 Or YOY4 < NetMarginMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY4 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    Selection.NumberFormat = "0.0%"
    If NetMargin3 < 0 Or YOY3 < NetMarginMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY4 - YOY3) > SlowNetMarginGrowth Or YOY3 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    Selection.NumberFormat = "0.0%"
    If NetMargin2 < 0 Or YOY2 < NetMarginMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY3 - YOY2) > SlowNetMarginGrowth Or YOY2 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    Selection.NumberFormat = "0.0%"
    If NetMargin1 < 0 Or YOY1 < NetMarginMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY2 - YOY1) > SlowNetMarginGrowth Or YOY1 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function

Sub ROE()

    Dim ErrorNum As Integer
    
    On Error GoTo ErrorHandler

'   name Revenue cell
    Range("B10").Name = "ROE"
    Range("10:10").Name = "ROERow"
    
'   write "Revenue" text
    Range("ROE").HorizontalAlignment = xlLeft
    Range("ROE") = "ROE"
    
    Range("ROERow").NumberFormat = "0.0%"
    
'   populate ROE information
    ErrorNum = 1
    ROE1 = NetIncome1 / Equity1
    If ROE1 >= ROERequirement Then
        Range("ROE").Offset(0, 1).Font.ColorIndex = GreenFont
    ElseIf ROE1 >= 0 Then
        Range("ROE").Offset(0, 1).Font.ColorIndex = OrangeFont
    Else
        Range("ROE").Offset(0, 1).Font.ColorIndex = RedFont
    End If
    Range("ROE").Offset(0, 1) = ROE1
    
    ErrorNum = 2
    ROE2 = NetIncome2 / Equity2
    If ROE2 >= ROERequirement Then
        Range("ROE").Offset(0, 2).Font.ColorIndex = GreenFont
    ElseIf ROE2 >= 0 Then
        Range("ROE").Offset(0, 2).Font.ColorIndex = OrangeFont
    Else
        Range("ROE").Offset(0, 2).Font.ColorIndex = RedFont
    End If
    Range("ROE").Offset(0, 2) = ROE2
    
    ErrorNum = 3
    ROE3 = NetIncome3 / Equity3
    If ROE3 >= ROERequirement Then
        Range("ROE").Offset(0, 3).Font.ColorIndex = GreenFont
    ElseIf ROE3 >= 0 Then
        Range("ROE").Offset(0, 3).Font.ColorIndex = OrangeFont
    Else
        Range("ROE").Offset(0, 3).Font.ColorIndex = RedFont
    End If
    Range("ROE").Offset(0, 3) = ROE3
    
    ErrorNum = 4
    ROE4 = NetIncome4 / Equity4
    If ROE4 >= ROERequirement Then
        Range("ROE").Offset(0, 4).Font.ColorIndex = GreenFont
    ElseIf ROE4 >= 0 Then
        Range("ROE").Offset(0, 4).Font.ColorIndex = OrangeFont
    Else
        Range("ROE").Offset(0, 4).Font.ColorIndex = RedFont
    End If
    Range("ROE").Offset(0, 4) = ROE4
    
    ErrorNum = 5
    ROE5 = NetIncome5 / Equity5
    If ROE5 >= ROERequirement Then
        Range("ROE").Offset(0, 5).Font.ColorIndex = GreenFont
    ElseIf ROE5 >= 0 Then
        Range("ROE").Offset(0, 5).Font.ColorIndex = OrangeFont
    Else
        Range("ROE").Offset(0, 5).Font.ColorIndex = RedFont
    End If
    Range("ROE").Offset(0, 5) = ROE5
    
    Range("ROE").AddComment
    Range("ROE").Comment.Visible = False
    Range("ROE").Comment.Text Text:="net income/equity" & Chr(10) & _
                "to increase earnings"
    Range("ROE").Comment.Shape.TextFrame.AutoSize = True
    
'   calculate YOY ROE growth
    ROEYOY
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case 1
            ROE1 = 0
            Range("ROE").Offset(0, 1) = ROE1
        Case 2
            ROE2 = 0
            Range("ROE").Offset(0, 2) = ROE2
        Case 3
            ROE3 = 0
            Range("ROE").Offset(0, 3) = ROE3
        Case 4
            ROE4 = 0
            Range("ROE").Offset(0, 4) = ROE4
        Case 5
            ROE5 = 0
            Range("ROE").Offset(0, 5) = ROE5
   End Select
   
   ROEYOY

End Sub

Sub ROEYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B11").Name = "YOYGrowth"
    Range("11:11").Name = "YOYRow"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(ROE1, ROE2)
    YOY2 = YOYGrowth(ROE2, ROE3)
    YOY3 = YOYGrowth(ROE3, ROE4)
    YOY4 = YOYGrowth(ROE4, ROE5)
    
    Call YOYROEEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYROEEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    Selection.NumberFormat = "0.0%"
    If ROE4 < 0 Or YOY4 < ROEMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY4 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    Selection.NumberFormat = "0.0%"
    If ROE3 < 0 Or YOY3 < ROEMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY4 - YOY3) > SlowROEGrowth Or YOY3 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    Selection.NumberFormat = "0.0%"
    If ROE2 < 0 Or YOY2 < ROEMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY3 - YOY2) > SlowROEGrowth Or YOY2 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    Selection.NumberFormat = "0.0%"
    If ROE1 < 0 Or YOY1 < ROEMaxDecline Then
        Selection.Font.ColorIndex = RedFont
    ElseIf (YOY2 - YOY1) > SlowROEGrowth Or YOY1 < 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function
