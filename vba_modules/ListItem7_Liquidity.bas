Attribute VB_Name = "ListItem7_Liquidity"
Option Explicit

Private dblQuickRatio(0 To 4) As Double

Private Const QUICK_RATIO_MIN = 1
Private ResultLiquidity As Result

'===============================================================
' Procedure:    EvaluateQuickRatio
'
' Description:  Display quick ratio information.
'               Call procedure to display YOY growth information
'               if quick ratio is greater than or equal to required min -> green font
'               else -> red font
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
' Rev History:  19ept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateQuickRatio()

    Dim ErrorNum As Years   'used to catch errors for each year of data
    
    On Error GoTo ErrorHandler

    ResultLiquidity = PASS
    
    Range("ListItemQuickRatio") = "Are debts covered?"
    Range("QuickRatio") = "Quick Ratio"
    
    Range("QuickRatio").AddComment
    Range("QuickRatio").Comment.Visible = False
    Range("QuickRatio").Comment.Text Text:="quick ratio = (current assets - inventory) / current liabilities" & Chr(10) & _
                "must be > 2 and not decreasing" & Chr(10) & _
                "better measure than current ratio which includes inventory and is thus higher"
    Range("QuickRatio").Comment.Shape.TextFrame.AutoSize = True
    
    ErrorNum = Year0
    
    dblQuickRatio(0) = (dblCurrentAssets(0) - dblInventory(0)) / dblCurrentLiabilities(0)
    If dblQuickRatio(0) >= QUICK_RATIO_MIN Then     'if quick ratio is greater than the required minimum
        Range("QuickRatio").Offset(0, 1).Font.ColorIndex = FONT_COLOR_GREEN
    Else                                            'if quick ratio is less than the required minimum
        Range("QuickRatio").Offset(0, 1).Font.ColorIndex = FONT_COLOR_RED
        ResultLiquidity = FAIL
    End If
    Range("QuickRatio").Offset(0, 1) = dblQuickRatio(0)
    
    ErrorNum = Year1
    dblQuickRatio(1) = (dblCurrentAssets(1) - dblInventory(1)) / dblCurrentLiabilities(1)
    If dblQuickRatio(1) >= QUICK_RATIO_MIN Then
        Range("QuickRatio").Offset(0, 2).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("QuickRatio").Offset(0, 2).Font.ColorIndex = FONT_COLOR_ORANGE
    End If
    Range("QuickRatio").Offset(0, 2) = dblQuickRatio(1)
    
    ErrorNum = Year2
    dblQuickRatio(2) = (dblCurrentAssets(2) - dblInventory(2)) / dblCurrentLiabilities(2)
    If dblQuickRatio(2) >= QUICK_RATIO_MIN Then
        Range("QuickRatio").Offset(0, 3).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("QuickRatio").Offset(0, 3).Font.ColorIndex = FONT_COLOR_ORANGE
    End If
    Range("QuickRatio").Offset(0, 3) = dblQuickRatio(2)
    
    ErrorNum = Year3
    dblQuickRatio(3) = (dblCurrentAssets(3) - dblInventory(3)) / dblCurrentLiabilities(3)
    If dblQuickRatio(3) >= QUICK_RATIO_MIN Then
        Range("QuickRatio").Offset(0, 4).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("QuickRatio").Offset(0, 4).Font.ColorIndex = FONT_COLOR_ORANGE
    End If
    Range("QuickRatio").Offset(0, 4) = dblQuickRatio(3)
    
    CalculateQuickRatioYOYGrowth
    
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case Year0
            dblQuickRatio(0) = 0
            Range("QuickRatio").Offset(0, 1) = dblQuickRatio(0)
        Case Year1
            dblQuickRatio(1) = 0
            Range("QuickRatio").Offset(0, 2) = dblQuickRatio(1)
        Case Year2
            dblQuickRatio(2) = 0
            Range("QuickRatio").Offset(0, 3) = dblQuickRatio(2)
        Case Year3
            dblQuickRatio(3) = 0
            Range("QuickRatio").Offset(0, 4) = dblQuickRatio(3)
   End Select
    
   CalculateQuickRatioYOYGrowth

End Sub

'===============================================================
' Procedure:    CalculateQuickRatioYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for quick ratio data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  19Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateQuickRatioYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double
    
    Range("QuickRatioYOYGrowth") = "YOY Growth (%)"

    'populate YOY growth information
    '(0) is most recent year
    dblYOYGrowth(0) = CalculateYOYGrowth(dblQuickRatio(0), dblQuickRatio(1))
    dblYOYGrowth(1) = CalculateYOYGrowth(dblQuickRatio(1), dblQuickRatio(2))
    dblYOYGrowth(2) = CalculateYOYGrowth(dblQuickRatio(2), dblQuickRatio(3))
    
    Call EvaluateQuickRatioYOYGrowth(Range("QuickRatioYOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2))
    
End Sub

'===============================================================
' Procedure:    EvaluateQuickRatioYOYGrowth
'
' Description:  Display YOY growth information.
'               for the most recent year
'                   if quick ratio is less than the min and decreasing -> red font
'                   else if quick ratio is decreasing -> orange font
'                   else quick ratio is increasing -> green font
'               for previous years
'                   if quick ratio is less than the min or decreasing -> orange font
'                   else quick ratio is increasing -> green font
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
' Rev History:  19Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateQuickRatioYOYGrowth(YOYGrowth As Range, YOY1, YOY2, YOY3)
    
    YOYGrowth.Offset(0, 3).Select
    If dblQuickRatio(2) < QUICK_RATIO_MIN Or YOY3 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If dblQuickRatio(1) < QUICK_RATIO_MIN Or YOY2 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If dblQuickRatio(0) < QUICK_RATIO_MIN And YOY1 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultLiquidity = FAIL
    ElseIf YOY1 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
    CheckLiquidityPassFail
    
End Function

'===============================================================
' Procedure:    CheckLiquidityPassFail
'
' Description:  Display check or x mark if the leverage
'               passes or fails the criteria
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  14Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CheckLiquidityPassFail()

    If ResultLiquidity = PASS Then
        Range("LiquidityCheck") = CHECK_MARK
        Range("LiquidityCheck").Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("LiquidityCheck") = X_MARK
        Range("LiquidityCheck").Font.ColorIndex = FONT_COLOR_RED
    End If

End Sub


