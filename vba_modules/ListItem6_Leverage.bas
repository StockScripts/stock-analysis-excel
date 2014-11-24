Attribute VB_Name = "ListItem6_Leverage"
Option Explicit

Private dblLeverageRatio(0 To 4) As Double
Private dblDebtToEquity(0 To 4) As Double

Private Const DEBT_TO_EQUITY_MAX = 0.4
Private Const LEVERAGE_RATIO_MAX = 2
Private ResultLeverage As Result

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
'               if leverage ratio is less than max -> green font
'               else -> red font
'
'               catch divide by 0 errors
'               ErrorNum serves as markers to indicate which
'               year data generates the error
'               -> set value to 0 if error
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

    Dim ErrorNum As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'financial leverage = assets / equity = (equity + liabilities) / equity
    '                   = 1 + (liabilities/equity)
    ErrorNum = 0
    dblLeverageRatio(0) = dblLiabilities(0) / dblEquity(0)
    If dblLeverageRatio(0) <= LEVERAGE_RATIO_MAX Then           'if financial leverage is less than max
        Range("LeverageRatio").Offset(0, 1).Font.ColorIndex = FONT_COLOR_GREEN
    Else                                                            'if financial leverage is greater than max
        Range("LeverageRatio").Offset(0, 1).Font.ColorIndex = FONT_COLOR_RED
        ResultLeverage = FAIL
    End If
    Range("LeverageRatio").Offset(0, 1) = dblLeverageRatio(0)
    
    For i = 1 To 3
        ErrorNum = i
        dblLeverageRatio(i) = dblLiabilities(i) / dblEquity(i)
        If dblLeverageRatio(i) <= LEVERAGE_RATIO_MAX Then
            Range("LeverageRatio").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_GREEN
        Else
            Range("LeverageRatio").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_ORANGE     'warning for past years
        End If                                                                              'only recent year should be looked at
        Range("LeverageRatio").Offset(0, i + 1) = dblLeverageRatio(i)
    Next i

    CalculateLeverageRatioYOYGrowth
    
    Exit Sub
    
ErrorHandler:
    Select Case ErrorNum
        Case 0
            dblLeverageRatio(0) = 0
            Range("LeverageRatio").Offset(0, 1) = dblLeverageRatio(0)
        Case 1
            dblLeverageRatio(1) = 0
            Range("LeverageRatio").Offset(0, 2) = dblLeverageRatio(1)
        Case 2
            dblLeverageRatio(2) = 0
            Range("LeverageRatio").Offset(0, 3) = dblLeverageRatio(2)
        Case 3
            dblLeverageRatio(3) = 0
            Range("LeverageRatio").Offset(0, 4) = dblLeverageRatio(3)
   End Select
   
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
    
    Call EvaluateLeverageRatioYOYGrowth(Range("LeverageRatioYOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2))
    
End Sub

'===============================================================
' Procedure:    EvaluateLeverageRatioYOYGrowth
'
' Description:  Display YOY growth information.
'               for the most recent year
'                   if leverage ratio is greater than max and increasing -> red font
'                   else if leverage ratio is increasing -> orange font
'                   else leverage ratio is decreasing -> green font
'               for previous years
'                   if leverage ratio is greater than max or increasing -> orange font
'                   else leverage ratio is decreasing -> green font
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   YOYGrowth As Range -> first cell of net margin YOY growth
'               YOY1, YOY2, YOY3 -> YOY growth values
'                                   (YOY1 is most recent year)
'
' Returns:      N/A
'
' Rev History:  13Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateLeverageRatioYOYGrowth(YOYGrowth As Range, YOY1, YOY2, YOY3)
    
    YOYGrowth.Offset(0, 3).Select
    If dblLeverageRatio(2) > LEVERAGE_RATIO_MAX Or YOY3 > 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If dblLeverageRatio(1) > LEVERAGE_RATIO_MAX Or YOY2 > 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If dblLeverageRatio(0) > LEVERAGE_RATIO_MAX And YOY1 > 0 Then    'if debt to equity is greater than max and increasing
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultLeverage = FAIL
    ElseIf YOY1 > 0 Then                                            'if debt to equity is increasing
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else                                                            'debt to equity is stable or decreasing
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function

'===============================================================
' Procedure:    EvaluateDebtToEquity
'
' Description:  Display debt to equity information.
'               Call procedure to display YOY growth information
'               if debt to equity is less than max -> green font
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
' Rev History:  18Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateDebtToEquity()

    Dim ErrorNum As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ErrorNum = 0
    dblDebtToEquity(0) = dblTotalDebt(0) / dblEquity(0)
    If dblDebtToEquity(0) <= LEVERAGE_RATIO_MAX Then       'if financial leverage is less than max
        Range("DebtToEquity").Offset(0, 1).Font.ColorIndex = FONT_COLOR_GREEN
    Else                                                            'if financial leverage is greater than max
        Range("DebtToEquity").Offset(0, 1).Font.ColorIndex = FONT_COLOR_RED
        ResultLeverage = FAIL
    End If
    Range("DebtToEquity").Offset(0, 1) = dblDebtToEquity(0)
    
    For i = 1 To 3
        ErrorNum = i
        dblDebtToEquity(i) = dblTotalDebt(i) / dblEquity(i)
        If dblDebtToEquity(i) <= LEVERAGE_RATIO_MAX Then
            Range("DebtToEquity").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_GREEN
        Else
            Range("DebtToEquity").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_ORANGE     'warning for past years
        End If                                                                                  'only recent year should be looked at
        Range("DebtToEquity").Offset(0, i + 1) = dblDebtToEquity(i)
    Next i

    CalculateDebtToEquityYOYGrowth
    
    Exit Sub
    
ErrorHandler:
    Select Case ErrorNum
        Case 0
            dblDebtToEquity(0) = 0
            Range("DebtToEquity").Offset(0, 1) = dblDebtToEquity(0)
        Case 1
            dblDebtToEquity(1) = 0
            Range("DebtToEquity").Offset(0, 2) = dblDebtToEquity(1)
        Case 2
            dblDebtToEquity(2) = 0
            Range("DebtToEquity").Offset(0, 3) = dblDebtToEquity(2)
        Case 3
            dblDebtToEquity(3) = 0
            Range("DebtToEquity").Offset(0, 4) = dblDebtToEquity(3)
   End Select
   
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
    
    Call EvaluateDebtToEquityYOYGrowth(Range("DebttoEquityYOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2))
    
End Sub

'===============================================================
' Procedure:    EvaluateDebtToEquityYOYGrowth
'
' Description:  Display YOY growth information.
'               for the most recent year
'                   if debt to equity is greater than max and increasing -> red font
'                   else if debt to equity is increasing -> orange font
'                   else debt to equity is decreasing -> green font
'               for previous years
'                   if debt to equity is greater than max or increasing -> orange font
'                   else debt is decreasing -> green font
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
' Rev History:  18Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateDebtToEquityYOYGrowth(YOYGrowth As Range, YOY1, YOY2, YOY3)
    
    YOYGrowth.Offset(0, 3).Select
    If dblDebtToEquity(2) > DEBT_TO_EQUITY_MAX Or YOY3 > 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If dblDebtToEquity(1) > DEBT_TO_EQUITY_MAX Or YOY2 > 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If dblDebtToEquity(0) > DEBT_TO_EQUITY_MAX And YOY1 > 0 Then    'if debt to equity is greater than max and increasing
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultLeverage = FAIL
    ElseIf YOY1 > 0 Then                                            'if debt to equity is increasing
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else                                                            'debt to equity is stable or decreasing
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
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

