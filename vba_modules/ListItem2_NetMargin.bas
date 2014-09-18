Attribute VB_Name = "ListItem2_NetMargin"
Option Explicit

Private dblNetMargin(0 To 4) As Double

'===============================================================
' Procedure:    EvaluateNetMargin
'
' Description:  Display net margin information.
'               Call procedure to display YOY growth information
'               if net margin is negative -> red font
'               else -> green font
'
'               catch divide by 0 errors
'               ErrorNum serves as markers to indicate which
'               year data generates the error
'               -> set growth to 0 if error
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
Sub EvaluateNetMargin()
    
    Dim ErrorNum As Years
    
    On Error GoTo ErrorHandler
    
    Range("A5").Font.Bold = True
    Range("A5") = "Are profits increasing?"
    
    'name Net Margin cell
    Range("B6").Name = "NetMargin"
    Range("6:6").Name = "NetMarginRow"
    
    'write "Net Margin" text
    Range("NetMargin").HorizontalAlignment = xlLeft
    Range("NetMargin") = "Net Margin"
    
    Range("NetMarginRow").NumberFormat = "0.0%"
    
    Range("NetMargin").AddComment
    Range("NetMargin").Comment.Visible = False
    Range("NetMargin").Comment.Text Text:="Net Profit Margin = Net Income / Revenue" & Chr(10) & _
                "Net profit margin demonstrates how good a company is at converting revenue into profits." & Chr(10) & _
                "These values should be increasing or stable. Net margin growth along with revenue growth leads to increased earnings." & Chr(10) & _
                "Net Income = Revenue x Profit Margin"

    Range("NetMargin").Comment.Shape.TextFrame.AutoSize = True
    
    'catch divide by 0
    ErrorNum = Year0
    
    'net margin = net income / revenue
    dblNetMargin(0) = dblNetIncome(0) / dblRevenue(0)
    
    If dblNetMargin(0) > 0 Then     'if net margin is positive
        Range("NetMargin").Offset(0, 1).Font.ColorIndex = FONT_COLOR_GREEN
    Else                            'if net margin is 0 or negative
        Range("NetMargin").Offset(0, 1).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("NetMargin").Offset(0, 1) = dblNetMargin(0)
    
    ErrorNum = Year1
    
    dblNetMargin(1) = dblNetIncome(1) / dblRevenue(1)
    If dblNetMargin(1) > 0 Then
        Range("NetMargin").Offset(0, 2).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("NetMargin").Offset(0, 2).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("NetMargin").Offset(0, 2) = dblNetMargin(1)
    
    ErrorNum = Year2
    
    dblNetMargin(2) = dblNetIncome(2) / dblRevenue(2)
    If dblNetMargin(2) > 0 Then
        Range("NetMargin").Offset(0, 3).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("NetMargin").Offset(0, 3).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("NetMargin").Offset(0, 3) = dblNetMargin(2)
    
    ErrorNum = Year3
    
    dblNetMargin(3) = dblNetIncome(3) / dblRevenue(3)
    If dblNetMargin(3) > 0 Then
        Range("NetMargin").Offset(0, 4).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("NetMargin").Offset(0, 4).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("NetMargin").Offset(0, 4) = dblNetMargin(3)
    
    ErrorNum = Year4
    
    dblNetMargin(4) = dblNetIncome(4) / dblRevenue(4)
    If dblNetMargin(4) > 0 Then
        Range("NetMargin").Offset(0, 5).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("NetMargin").Offset(0, 5).Font.ColorIndex = FONT_COLOR_RED
    End If
    Range("NetMargin").Offset(0, 5) = dblNetMargin(4)
    
    CalculateNetMarginYOYGrowth
    
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case Year0
            dblNetMargin(0) = 0
            Range("NetMargin").Offset(0, 1) = dblNetMargin(0)
        Case Year1
            dblNetMargin(1) = 0
            Range("NetMargin").Offset(0, 2) = dblNetMargin(1)
        Case Year2
            dblNetMargin(2) = 0
            Range("NetMargin").Offset(0, 3) = dblNetMargin(2)
        Case Year3
            dblNetMargin(3) = 0
            Range("NetMargin").Offset(0, 4) = dblNetMargin(3)
        Case Year4
            dblNetMargin(4) = 0
            Range("NetMargin").Offset(0, 5) = dblNetMargin(4)
   End Select
   
   CalculateNetMarginYOYGrowth
    
End Sub

'===============================================================
' Procedure:    CalculateNetMarginYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for net margin data. Format cells.
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
Sub CalculateNetMarginYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double
    
'   name YOY cell
    Range("B7").Name = "YOYGrowth"
    Range("7:7").Name = "YOYRow"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    Range("YOYRow").NumberFormat = "0.0%"
    
    'populate YOY growth information
    '(0) is most recent year
    dblYOYGrowth(0) = CalculateYOYGrowth(dblNetMargin(0), dblNetMargin(1))
    dblYOYGrowth(1) = CalculateYOYGrowth(dblNetMargin(1), dblNetMargin(2))
    dblYOYGrowth(2) = CalculateYOYGrowth(dblNetMargin(2), dblNetMargin(3))
    dblYOYGrowth(3) = CalculateYOYGrowth(dblNetMargin(3), dblNetMargin(4))
    
    Call EvaluateNetMarginYOYGrowth(Range("YOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2), dblYOYGrowth(3))
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

'===============================================================
' Procedure:    EvaluateNetMarginYOYGrowth
'
' Description:  Display YOY growth information.
'               if net margin decreases -> red font
'               else net margin growth is positive -> green font
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   YOYGrowth As Range -> first cell of net margin YOY growth
'               YOY1, YOY2, YOY3, YOY4 -> YOY growth values
'                                         (YOY1 is most recent year)
'
' Returns:      N/A
'
' Rev History:  17Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateNetMarginYOYGrowth(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    Selection.NumberFormat = "0.0%"
    If dblNetMargin(3) < 0 Or YOY4 < 0 Then     'if net margin is negative or net margin decreases
        Selection.Font.ColorIndex = FONT_COLOR_RED
    Else                                        'net margin is stable or increasing
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    Selection.NumberFormat = "0.0%"
    If dblNetMargin(2) < 0 Or YOY3 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    Selection.NumberFormat = "0.0%"
    If dblNetMargin(1) < 0 Or YOY2 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    Selection.NumberFormat = "0.0%"
    If dblNetMargin(0) < 0 Or YOY1 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function

