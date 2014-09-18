Attribute VB_Name = "ListItem3_Earnings"
Option Explicit

Private Const EPS_GROWTH_MIN = 0.1  'EPS must grow by 10% each year

'===============================================================
' Procedure:    EvaluateEPS
'
' Description:  Display EPS information.
'               Call procedure to display YOY growth information
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  17Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateEPS()

    Range("A8").Font.Bold = True
    Range("A8") = "Are earnings increasing?"
    
'   name EPS cell
    Range("B9").Name = "DilutedEPS"
    
'   write "Diluted EPS" text
    Range("DilutedEPS").HorizontalAlignment = xlLeft
    Range("DilutedEPS") = "Diluted EPS"
    
'   populate EPS information
    If dblEPS(0) > 0 Then
        Range("DilutedEPS").Offset(0, 1).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("DilutedEPS").Offset(0, 1).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("DilutedEPS").Offset(0, 1) = dblEPS(0)
    
    If dblEPS(1) > 0 Then
        Range("DilutedEPS").Offset(0, 2).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("DilutedEPS").Offset(0, 2).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("DilutedEPS").Offset(0, 2) = dblEPS(1)

    If dblEPS(2) > 0 Then
        Range("DilutedEPS").Offset(0, 3).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("DilutedEPS").Offset(0, 3).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("DilutedEPS").Offset(0, 3) = dblEPS(2)
    
    If dblEPS(3) > 0 Then
        Range("DilutedEPS").Offset(0, 4).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("DilutedEPS").Offset(0, 4).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("DilutedEPS").Offset(0, 4) = dblEPS(3)
    
    If dblEPS(4) > 0 Then
        Range("DilutedEPS").Offset(0, 5).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("DilutedEPS").Offset(0, 5).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("DilutedEPS").Offset(0, 5) = dblEPS(4)
    
    Range("DilutedEPS").AddComment
    Range("DilutedEPS").Comment.Visible = False
    Range("DilutedEPS").Comment.Text Text:="EPS = Net Income / Shares Outstanding" & Chr(10) & _
                "EPS ultimately drive stock prices and increasing earnings generally moves the stock price up." & Chr(10) & _
                "Earnings or Net Income should follow revenue growth if the profit margin remains the same." & Chr(10) & _
                "If the Net Profit Margin increases, Net Income and EPS should grow faster than Revenue."
    Range("DilutedEPS").Comment.Shape.TextFrame.AutoSize = True
    
    CalculateEPSYOYGrowth

End Sub

'===============================================================
' Procedure:    CalculateEPSYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for EPS data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  17Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateEPSYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double
    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B10").Name = "YOYGrowth"
    Range("10:10").Name = "YOYRow"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    Range("YOYRow").NumberFormat = "0.0%"
    
    'populate YOY growth information
    '(0) is most recent year
    dblYOYGrowth(0) = CalculateYOYGrowth(dblEPS(0), dblEPS(1))
    dblYOYGrowth(1) = CalculateYOYGrowth(dblEPS(1), dblEPS(2))
    dblYOYGrowth(2) = CalculateYOYGrowth(dblEPS(2), dblEPS(3))
    dblYOYGrowth(3) = CalculateYOYGrowth(dblEPS(3), dblEPS(4))
    
    Call EvaluateEPSYOYGrowth(Range("YOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2), dblYOYGrowth(3))
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

'===============================================================
' Procedure:    EvaluateEPSYOYGrowth
'
' Description:  Display YOY growth information.
'               if EPS decreases -> red font
'               if EPS growth is < EPS_GROWTH_MIN -> orange font
'               else if EPS growth >= EPS_GROWTH_MIN -> green font
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   YOYGrowth As Range -> first cell of revenue YOY growth
'               YOY1, YOY2, YOY3, YOY4 -> YOY growth values
'                                         (YOY1 is most recent year)
'
' Returns:      N/A
'
' Rev History:  17Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateEPSYOYGrowth(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    If dblEPS(3) < 0 Or YOY4 < 0 Then               'if EPS is negative or decreases
        Selection.Font.ColorIndex = FONT_COLOR_RED
    ElseIf YOY4 < EPS_GROWTH_MIN Then               'if EPS growth is less than required
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else                                            'if EPS growth is greater than required
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    If dblEPS(2) < 0 Or YOY3 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
    ElseIf YOY3 < EPS_GROWTH_MIN Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If dblEPS(1) < 0 Or YOY2 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
    ElseIf YOY2 < EPS_GROWTH_MIN Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If dblEPS(0) < 0 Or YOY1 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
    ElseIf YOY1 < EPS_GROWTH_MIN Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function


