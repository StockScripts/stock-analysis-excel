Attribute VB_Name = "ListItem6_Leverage"
Option Explicit

Private dblLeverageRatio(0 To 4) As Double
Private dblDebtToEquity(0 To 4) As Double

Private Const DEBT_TO_EQUITY_MAX = 0.4
Private Const LEVERAGE_RATIO_MAX = 2
Private ResultLeverage As Result
Private Const LEVERAGE_SCORE_MAX = 4
Private Const LEVERAGE_SCORE_WEIGHT = 3
Private ScoreLeverage As Integer

'===============================================================
' Procedure:    EvaluateFinancialLeverage
'
' Description:  Call procedures to evaluate leverage ratio and
'               debt
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  13Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateFinancialLeverage()

    ResultLeverage = PASS
    
    ScoreLeverage = 0
    
    EvaluateLeverageRatio
    EvaluateDebtToEquity
    
    DisplayLeverageInfo
    
    CheckLeveragePassFail

End Sub

'===============================================================
' Procedure:    DisplayLeverageInfo
'
' Description:  Comment box information for leverage
'               - leverage requirements
'               - leverage information information
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
Sub DisplayLeverageInfo()
    
    Dim dblAssetsYOYGrowth(0 To 3) As Double
    Dim strAssetsYOYGrowth(0 To 3) As String
    
    Dim dblLiabilitiesYOYGrowth(0 To 3) As Double
    Dim strLiabilitiesYOYGrowth(0 To 3) As String
    
    Dim dblTotalDebtYOYGrowth(0 To 3) As Double
    Dim strTotalDebtYOYGrowth(0 To 3) As String
    
    Dim dblEquityYOYGrowth(0 To 2) As Double
    Dim strEquityYOYGrowth(0 To 2) As String
    
    Dim i As Integer
    
    Range("ListItemFinancialLeverage") = "Is it leveraged?"
    Range("LeverageRatio") = "Leverage Ratio"
    Range("DebtToEquity") = "Debt To Equity"

    With Range("ListItemFinancialLeverage")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="What is it:" & Chr(10) & _
                "   Financial leverage is the use of borrowed money to finance assets." & Chr(10) & _
                "   One measure of leverage is to divide total liabilities by equity." & Chr(10) & _
                "   A leverage ratio of 2 means for every dollar of equity, the company has 2 dollars of liability." & Chr(10) & _
                "   Total debt to equity is another measure of financial leverage." & Chr(10) & _
                "Why is it important:" & Chr(10) & _
                "   Increased leverage increases potential profitability, but also potential risk." & Chr(10) & _
                "   A high debt to equity ratio can result in volatile earnings due to additional interenst expense." & Chr(10) & _
                "What to look for:" & Chr(10) & _
                "   The recent year leverage ratio should not exceed 2." & Chr(10) & _
                "   The recent year debt to equity value should not exceed 40%." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   Increasing ROE may be due to increasing leverage ratio."
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    'calculate YOY growth
    For i = 0 To (iYearsAvailableIncome - 2)
        dblAssetsYOYGrowth(i) = CalculateYOYGrowth(dblAssets(i), dblAssets(i + 1))
        strAssetsYOYGrowth(i) = Format(dblAssetsYOYGrowth(i), "0.0%")
        
        dblLiabilitiesYOYGrowth(i) = CalculateYOYGrowth(dblLiabilities(i), dblLiabilities(i + 1))
        strLiabilitiesYOYGrowth(i) = Format(dblLiabilitiesYOYGrowth(i), "0.0%")
        
        dblTotalDebtYOYGrowth(i) = CalculateYOYGrowth(dblTotalDebt(i), dblTotalDebt(i + 1))
        strTotalDebtYOYGrowth(i) = Format(dblTotalDebtYOYGrowth(i), "0.0%")
        
        dblEquityYOYGrowth(i) = CalculateYOYGrowth(dblEquity(i), dblEquity(i + 1))
        strEquityYOYGrowth(i) = Format(dblEquityYOYGrowth(i), "0.0%")
    Next i
    
    With Range("LeverageRatio")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="Assets/Equity = (Equity + Liabilities)/Equity " & Chr(10) & _
                "                   = 1 + (Liabilities/Equity)" & Chr(10) & _
                "" & Chr(10) & _
                "Leverage Ratio = Liabilities/Equity" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Total Assets" & "               " & dblAssets(0) & "      " & dblAssets(1) & "      " & dblAssets(2) & "      " & dblAssets(3) & Chr(10) & _
                "YOY Total Debt Growth     " & strAssetsYOYGrowth(0) & "     " & strAssetsYOYGrowth(1) & "     " & strAssetsYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Total Liabilities" & "                " & dblLiabilities(0) & "      " & dblLiabilities(1) & "      " & dblLiabilities(2) & "      " & dblLiabilities(3) & Chr(10) & _
                "YOY Total Liabilities Growth     " & strLiabilitiesYOYGrowth(0) & "     " & strLiabilitiesYOYGrowth(1) & "     " & strLiabilitiesYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Equity              " & dblEquity(0) & "     " & dblEquity(1) & "     " & dblEquity(2) & "     " & dblEquity(3) & Chr(10) & _
                "YOY Equity Growth   " & strEquityYOYGrowth(0) & "     " & strEquityYOYGrowth(1) & "     " & strEquityYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    With Range("DebtToEquity")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="Debt to Equity = Total Debt/Equity " & Chr(10) & _
                "" & Chr(10) & _
                "YOY Total Debt" & "                " & dblTotalDebt(0) & "      " & dblTotalDebt(1) & "      " & dblTotalDebt(2) & "      " & dblTotalDebt(3) & Chr(10) & _
                "YOY Total Debt Growth     " & strTotalDebtYOYGrowth(0) & "     " & strTotalDebtYOYGrowth(1) & "     " & strTotalDebtYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Equity              " & dblEquity(0) & "     " & dblEquity(1) & "     " & dblEquity(2) & "     " & dblEquity(3) & Chr(10) & _
                "YOY Equity Growth   " & strEquityYOYGrowth(0) & "     " & strEquityYOYGrowth(1) & "     " & strEquityYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With

End Sub

'===============================================================
' Procedure:    EvaluateLeverageRatio
'
' Description:  Display leverage ratio information.
'               Call procedure to display YOY growth information
'               if recent year leverage ratio is less than max -> pass
'               else -> fail
'
'               if previous years leverage ratio is less than max -> pass
'               else -> warning
'
'               catch divide by 0 errors
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  13Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateLeverageRatio()

    Dim i As Integer
    
    On Error Resume Next
    
    'financial leverage = assets / equity = (equity + liabilities) / equity
    '                   = 1 + (liabilities/equity)
    Range("LeverageRatio").Offset(0, 1).Select
    dblLeverageRatio(0) = dblLiabilities(0) / dblEquity(0)
    
    If Err Then
        Selection.HorizontalAlignment = xlCenter
        Selection.Value = STR_NO_DATA
        Err.Clear
    Else
        If dblLeverageRatio(0) <= LEVERAGE_RATIO_MAX Then       'if financial leverage is less than max
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreLeverage = ScoreLeverage + (LEVERAGE_SCORE_MAX - i)
        Else                                                    'if financial leverage is greater than max
            Selection.Font.ColorIndex = FONT_COLOR_RED
            ResultLeverage = FAIL
        End If
        Selection.Value = dblLeverageRatio(0)
    End If
    
    For i = 1 To 3
        dblLeverageRatio(i) = dblLiabilities(i) / dblEquity(i)
        Range("LeverageRatio").Offset(0, i + 1).Select
        If Err Then
            Selection.HorizontalAlignment = xlCenter
            Selection.Value = STR_NO_DATA
            Err.Clear
        Else
            If dblLeverageRatio(i) <= LEVERAGE_RATIO_MAX Then
                Selection.Font.ColorIndex = FONT_COLOR_GREEN
                ScoreLeverage = ScoreLeverage + (LEVERAGE_SCORE_MAX - i)
            Else
               Selection.Font.ColorIndex = FONT_COLOR_ORANGE     'warning for past years
            End If                                                                              'only recent year should be looked at
            Selection.Value = dblLeverageRatio(i)
        End If
    Next i

    CalculateLeverageRatioYOYGrowth

End Sub

'===============================================================
' Procedure:    CalculateLeverageRatioYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for leverage ratio data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  18Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateLeverageRatioYOYGrowth()

    Dim dblYOYGrowth(0 To 2) As Double
    Dim i As Integer

    Range("LeverageRatioYOYGrowth") = "YOY Growth (%)"
    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To 2
        dblYOYGrowth(i) = CalculateYOYGrowth(dblLeverageRatio(i), dblLeverageRatio(i + 1))
    Next i
    
    Call EvaluateLeverageRatioYOYGrowth(Range("LeverageRatioYOYGrowth"), dblYOYGrowth)
    
End Sub

'===============================================================
' Procedure:    EvaluateLeverageRatioYOYGrowth
'
' Description:  Display YOY growth information.
'               for the most recent year
'                   if leverage ratio is greater than max -> fail
'                   else if leverage ratio is increasing -> warning
'                   else leverage ratio is decreasing -> pass
'               for previous years
'                   if leverage ratio is greater than max or increasing -> warning
'                   else leverage ratio is decreasing -> pass
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   YOYGrowth As Range -> first cell of net margin YOY growth
'               YOY array -> YOY growth values
'                            YOY(0) is most recent year
'
' Returns:      N/A
'
' Rev History:  13Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateLeverageRatioYOYGrowth(YOYGrowth As Range, YOY() As Double)
    
    Dim i As Integer
    
    YOYGrowth.Offset(0, 1).Select
    If dblLeverageRatio(0) > LEVERAGE_RATIO_MAX Then        'if debt to equity is greater than max
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultLeverage = FAIL
    ElseIf YOY(0) > 0 Then                                  'if debt to equity is increasing
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else                                                    'debt to equity is stable or decreasing
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
        ScoreLeverage = ScoreLeverage + (LEVERAGE_SCORE_MAX - i)
    End If
    YOYGrowth.Offset(0, 1) = YOY(0)
    
    For i = 1 To (iYearsAvailableIncome - 2)
        YOYGrowth.Offset(0, i + 1).Select
        If dblLeverageRatio(i) > LEVERAGE_RATIO_MAX Or YOY(i) > 0 Then
            Selection.Font.ColorIndex = FONT_COLOR_ORANGE
        Else
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreLeverage = ScoreLeverage + (LEVERAGE_SCORE_MAX - i)
        End If
        Selection.Value = YOY(i)
    Next i
    
End Function

'===============================================================
' Procedure:    EvaluateDebtToEquity
'
' Description:  Display debt to equity information.
'               Call procedure to display YOY growth information
'               if recent year debt to equity is less than max -> pass
'               else -> fail
'
'               if previous years debt to equiyt is less than max -> pass
'               else -> warning
'
'               catch divide by 0 errors
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  18Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateDebtToEquity()

    Dim i As Integer
    
    On Error Resume Next
    
    dblDebtToEquity(0) = dblTotalDebt(0) / dblEquity(0)
    Range("DebtToEquity").Offset(0, 1).Select
    If Err Then
        Selection.HorizontalAlignment = xlCenter
        Selection.Value = STR_NO_DATA
        Err.Clear
    Else
        If dblDebtToEquity(0) <= DEBT_TO_EQUITY_MAX Then       'if financial leverage is less than max
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreLeverage = ScoreLeverage + (LEVERAGE_SCORE_MAX - i)
        Else                                                   'if financial leverage is greater than max
            Selection.Font.ColorIndex = FONT_COLOR_RED
            ResultLeverage = FAIL
        End If
        Selection.Value = dblDebtToEquity(0)
    End If
    
    For i = 1 To 3
        dblDebtToEquity(i) = dblTotalDebt(i) / dblEquity(i)
        Range("DebtToEquity").Offset(0, i + 1).Select
        If dblDebtToEquity(i) <= LEVERAGE_RATIO_MAX Then
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreLeverage = ScoreLeverage + (LEVERAGE_SCORE_MAX - i)
        Else
            Selection.Font.ColorIndex = FONT_COLOR_ORANGE     'warning for past years
        End If                                                'only recent year should be looked at
        Selection.Value = dblDebtToEquity(i)
    Next i

    CalculateDebtToEquityYOYGrowth

End Sub

'===============================================================
' Procedure:    CalculateDebtToEquityYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for debt to equity data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  18Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateDebtToEquityYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double
    Dim i As Integer
    
    Range("DebttoEquityYOYGrowth") = "YOY Growth (%)"

    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To 3
        dblYOYGrowth(i) = CalculateYOYGrowth(dblDebtToEquity(i), dblDebtToEquity(i + 1))
    Next i
    
    Call EvaluateDebtToEquityYOYGrowth(Range("DebttoEquityYOYGrowth"), dblYOYGrowth)
    
End Sub

'===============================================================
' Procedure:    EvaluateDebtToEquityYOYGrowth
'
' Description:  Display YOY growth information.
'               for the most recent year
'                   if debt to equity is greater than max and increasing -> fail
'                   else if debt to equity is increasing -> warning
'                   else debt to equity is decreasing -> pass
'               for previous years
'                   if debt to equity is greater than max or increasing -> warning
'                   else debt is decreasing -> pass
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   YOYGrowth As Range -> first cell of net margin YOY growth
'               YOY array -> YOY growth values
'                            YOY(0) is most recent year
'
' Returns:      N/A
'
' Rev History:  18Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateDebtToEquityYOYGrowth(YOYGrowth As Range, YOY() As Double)
    
    Dim i As Integer
    
    YOYGrowth.Offset(0, 1).Select
    If dblDebtToEquity(0) > DEBT_TO_EQUITY_MAX Then         'if debt to equity is greater than max
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultLeverage = FAIL
    ElseIf YOY(0) > 0 Then                                  'if debt to equity is increasing
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else                                                    'debt to equity is stable or decreasing
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
        ScoreLeverage = ScoreLeverage + (LEVERAGE_SCORE_MAX - i)
    End If
    YOYGrowth.Offset(0, 1) = YOY(0)
    
    For i = 1 To (iYearsAvailableIncome - 2)
        YOYGrowth.Offset(0, i + 1).Select
        If dblDebtToEquity(i) > DEBT_TO_EQUITY_MAX Or YOY(i) > 0 Then
            Selection.Font.ColorIndex = FONT_COLOR_ORANGE
        Else
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreLeverage = ScoreLeverage + (LEVERAGE_SCORE_MAX - i)
        End If
        Selection.Value = YOY(i)
    Next i
    
    ScoreLeverage = ScoreLeverage * LEVERAGE_SCORE_WEIGHT
    
    CheckLeveragePassFail
    LeverageScore
    
End Function

'===============================================================
' Procedure:    CheckLeveragePassFail
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
Sub CheckLeveragePassFail()

    If ResultLeverage = PASS Then
        Range("LeverageCheck") = CHECK_MARK
        Range("LeverageCheck").Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("LeverageCheck") = X_MARK
        Range("LeverageCheck").Font.ColorIndex = FONT_COLOR_RED
    End If

End Sub

'===============================================================
' Procedure:    LeverageScore
'
' Description:  Calculate score for leverage
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  10Dec15 by Janice Laset Parkerson
'               - Initial Version
'===============================================================

Sub LeverageScore()

    Range("LeverageScore") = ScoreLeverage

End Sub

