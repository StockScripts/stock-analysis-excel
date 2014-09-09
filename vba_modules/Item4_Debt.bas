Attribute VB_Name = "Item4_Debt"

Private Const MaxDebtIncrease = 0.3
Private Const DebtToEquityRequirement = 0.4

Public DebtToEquity1 As Double
Public DebtToEquity2 As Double
Public DebtToEquity3 As Double
Public DebtToEquity4 As Double
Public DebtToEquity5 As Double

Sub Debt()

    LongTermDebt
    DebtToEquity

End Sub

Sub LongTermDebt()

    Range("A15").Font.Bold = True
    Range("A15") = "Do they have a lot of debt?"
    
'   name LongTermDebt cell
    Range("B16").Name = "LongTermDebt"
    
'   write "LongTermDebt" text
    Range("LongTermDebt").HorizontalAlignment = xlLeft
    Range("LongTermDebt") = "Long Term Debt"
    
'   populate long term debt information
    Range("LongTermDebt").Offset(0, 1) = LongTermDebt1
    Range("LongTermDebt").Offset(0, 2) = LongTermDebt2
    Range("LongTermDebt").Offset(0, 3) = LongTermDebt3
    Range("LongTermDebt").Offset(0, 4) = LongTermDebt4
    Range("LongTermDebt").Offset(0, 5) = LongTermDebt5
    
'   calculate YOY long term debt growth
    LongTermDebtYOY

End Sub

Sub LongTermDebtYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B17").Name = "YOYGrowth"
    Range("17:17").Name = "YOYRow"
    
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
    
    YOY1 = YOYGrowth(LongTermDebt1, LongTermDebt2)
    YOY2 = YOYGrowth(LongTermDebt2, LongTermDebt3)
    YOY3 = YOYGrowth(LongTermDebt3, LongTermDebt4)
    YOY4 = YOYGrowth(LongTermDebt4, LongTermDebt5)
    
    Call YOYDebtEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYDebtEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    If YOY4 > MaxDebtIncrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY4 > 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    If YOY3 > MaxDebtIncrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY3 > 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If YOY2 > MaxDebtIncrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY2 > 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If YOY1 > MaxDebtIncrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY1 > 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function

Sub DebtToEquity()

    Dim ErrorNum As Integer
    
    On Error GoTo ErrorHandler

'   name Revenue cell
    Range("B18").Name = "DebtToEquity"
    Range("18:18").Name = "DebtToEquityRow"
    
'   write "Revenue" text
    Range("DebtToEquity").HorizontalAlignment = xlLeft
    Range("DebtToEquity") = "Debt To Equity"
    
    Range("DebtToEquityRow").NumberFormat = "0.0%"
    
'   populate ROE information
    ErrorNum = 1
    DebtToEquity1 = LongTermDebt1 / Equity1
    If DebtToEquity1 <= DebtToEquityRequirement Then
        Range("DebtToEquity").Offset(0, 1).Font.ColorIndex = GreenFont
    Else
        Range("DebtToEquity").Offset(0, 1).Font.ColorIndex = RedFont
    End If
    Range("DebtToEquity").Offset(0, 1) = DebtToEquity1
    
    ErrorNum = 2
    DebtToEquity2 = LongTermDebt2 / Equity2
    If DebtToEquity2 <= DebtToEquityRequirement Then
        Range("DebtToEquity").Offset(0, 2).Font.ColorIndex = GreenFont
    Else
        Range("DebtToEquity").Offset(0, 2).Font.ColorIndex = RedFont
    End If
    Range("DebtToEquity").Offset(0, 2) = DebtToEquity2
    
    ErrorNum = 3
    DebtToEquity3 = LongTermDebt3 / Equity3
    If DebtToEquity3 <= DebtToEquityRequirement Then
        Range("DebtToEquity").Offset(0, 3).Font.ColorIndex = GreenFont
    Else
        Range("DebtToEquity").Offset(0, 3).Font.ColorIndex = RedFont
    End If
    Range("DebtToEquity").Offset(0, 3) = DebtToEquity3
    
    ErrorNum = 4
    DebtToEquity4 = LongTermDebt4 / Equity4
    If DebtToEquity4 <= DebtToEquityRequirement Then
        Range("DebtToEquity").Offset(0, 4).Font.ColorIndex = GreenFont
    Else
        Range("DebtToEquity").Offset(0, 4).Font.ColorIndex = RedFont
    End If
    Range("DebtToEquity").Offset(0, 4) = DebtToEquity4
    
    ErrorNum = 5
    DebtToEquity5 = LongTermDebt5 / Equity5
    If DebtToEquity5 <= DebtToEquityRequirement Then
        Range("DebtToEquity").Offset(0, 5).Font.ColorIndex = GreenFont
    Else
        Range("DebtToEquity").Offset(0, 5).Font.ColorIndex = RedFont
    End If
    Range("DebtToEquity").Offset(0, 5) = DebtToEquity5

'   calculate YOY Debt To Equity growth
    DebtToEquityYOY
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case 1
            DebtToEquity1 = 0
            Range("DebtToEquity").Offset(0, 1).Select
            Selection.NumberFormat = "0.0%"
            Range("DebtToEquity").Offset(0, 1) = DebtToEquity1
        Case 2
            DebtToEquity2 = 0
            Range("DebtToEquity").Offset(0, 2).Select
            Selection.NumberFormat = "0.0%"
            Range("DebtToEquity").Offset(0, 2) = DebtToEquity2
        Case 3
            DebtToEquity3 = 0
            Range("DebtToEquity").Offset(0, 3).Select
            Selection.NumberFormat = "0.0%"
            Range("DebtToEquity").Offset(0, 3) = DebtToEquity3
        Case 4
            DebtToEquity4 = 0
            Range("DebtToEquity").Offset(0, 4) = DebtToEquity4
        Case 5
            DebtToEquity5 = 0
            Range("DebtToEquity").Offset(0, 5).Select
            Selection.NumberFormat = "0.0%"
            Range("DebtToEquity").Offset(0, 5) = DebtToEquity5
   End Select
   
   DebtToEquityYOY

End Sub

Sub DebtToEquityYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B19").Name = "YOYGrowth"
    Range("19:19").Name = "YOYRow"
    
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
    
    YOY1 = YOYGrowth(DebtToEquity1, DebtToEquity2)
    YOY2 = YOYGrowth(DebtToEquity2, DebtToEquity3)
    YOY3 = YOYGrowth(DebtToEquity3, DebtToEquity4)
    YOY4 = YOYGrowth(DebtToEquity4, DebtToEquity5)
    
    Call YOYDebtToEquityEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYDebtToEquityEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    If YOY4 > MaxDebtIncrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY4 > 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    If YOY3 > MaxDebtIncrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY3 > 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If YOY2 > MaxDebtIncrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY2 > 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If YOY1 > MaxDebtIncrease Then
        Selection.Font.ColorIndex = RedFont
    ElseIf YOY1 > 0 Then
        Selection.Font.ColorIndex = OrangeFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function


