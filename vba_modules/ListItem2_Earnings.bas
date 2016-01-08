Attribute VB_Name = "ListItem2_Earnings"
Option Explicit

Private Const EPS_GROWTH_MIN = 0.1  'EPS must grow by 10% each year
Private ResultEarnings As Result
Private Const EARNINGS_SCORE_MAX = 4
Private Const EARNINGS_SCORE_WEIGHT = 9
Public ScoreEarnings As Integer
Public Const MAX_EARNINGS_SCORE = 171

'===============================================================
' Procedure:    EvaluateEPS
'
' Description:  Display EPS information.
'               Call procedure to display YOY growth information
'               flag pass/fail for three most recent years
'               pass if positive, fail if negative
'               Scoring:
'               most recent year > 0
'                   add EARNINGS_SCORE_MAX
'                   subtract if earnings < 0
'               most recent year - 1 > 0
'                   add EARNINGS_SCORE_MAX
'                   subtract if earnings < 0
'               total score = score * EARNINGS_SCORE_WEIGHT
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
    
    Dim i As Integer
    
    ResultEarnings = PASS
    ScoreEarnings = 0
    
'   populate EPS information
    For i = 0 To (iYearsAvailableIncome - 1)
        If IsNumeric(vEPS(i)) Then
            If vEPS(i) > 0 Then
                Range("Earnings").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_GREEN
                ScoreEarnings = ScoreEarnings + (EARNINGS_SCORE_MAX - i)
            Else
                Range("Earnings").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_RED
                ResultEarnings = FAIL
                ScoreEarnings = ScoreEarnings - (EARNINGS_SCORE_MAX - i)
            End If
            Range("Earnings").Offset(0, i + 1) = vEPS(i)
        Else
            With Range("Earnings")
                .Offset(0, i + 1).HorizontalAlignment = xlCenter
                .Offset(0, i + 1) = STR_NO_DATA
            End With
        End If
    Next i
    
    DisplayEarningsInfo
    CalculateEPSYOYGrowth

End Sub

'===============================================================
' Procedure:    DisplayEarningsInfo
'
' Description:  Comment box information for EPS
'               - earnings requirements
'               - EPS information
'               - net income and YOY growth
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  28Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub DisplayEarningsInfo()

    Dim dblNetIncomeYOYGrowth(0 To 3) As Double
    Dim strNetIncomeYOYGrowth(0 To 3) As String
    
    Dim dblSharesYOYGrowth(0 To 3) As Double
    Dim strSharesYOYGrowth(0 To 3) As String
    
    Dim dblTaxRate(0 To 3) As Double
    Dim strTaxRate(0 To 3) As String
    Dim dblTaxRateYOYGrowth(0 To 3) As Double
    Dim strTaxRateYOYGrowth(0 To 3) As String
    
    Dim dblExpenseToSales(0 To 3) As Double
    Dim strExpenseToSales(0 To 3) As String
    Dim dblExpenseToSalesYOYGrowth(0 To 3) As Double
    Dim strExpenseToSalesYOYGrowth(0 To 3) As String
    
    Dim i As Integer
    
    On Error Resume Next
    
    Range("ListItemEarnings") = "Are earnings increasing?"
    Range("Earnings") = "Diluted EPS"
    
    With Range("ListItemEarnings")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="What is it:" & Chr(10) & _
                "   Earnings or EPS is the amount a company is earning for each share of its stock." & Chr(10) & _
                "   It is the net income per share of the company." & Chr(10) & _
                "Why is it important:" & Chr(10) & _
                "   EPS is a measure of profitability, and it generally drives stock prices." & Chr(10) & _
                "What to look for:" & Chr(10) & _
                "   EPS should increase by at least 10% every year." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   If EPS is increasing significantly faster than revenue, it could be due to" & Chr(10) & _
                "   a decrease in expenses or tax rate or the company could be buying back shares." & Chr(10) & _
                "   Constantly cutting costs may not be sustainable in the long term."
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    For i = 0 To (iYearsAvailableIncome - 1)
        dblExpenseToSales(i) = dblOperatingExpense(i) / dblRevenue(i)
        
        If Err Then
            dblExpenseToSales(i) = 0
            Err.Clear
        End If
        
        strExpenseToSales(i) = Format(dblExpenseToSales(i), "0.00")
        
        dblTaxRate(i) = 1 - (dblIncomeAfterTax(i) / dblIncomeBeforeTax(i))
        
        If Err Then
            dblTaxRate(i) = 0
            Err.Clear
        End If
        
        strTaxRate(i) = Format(dblTaxRate(i), "0.0%")
    Next i
    
    'calculate YOY growth
    For i = 0 To (iYearsAvailableIncome - 2)
        dblNetIncomeYOYGrowth(i) = CalculateYOYGrowth(dblNetIncome(i), dblNetIncome(i + 1))
        strNetIncomeYOYGrowth(i) = Format(dblNetIncomeYOYGrowth(i), "0.0%")
        
        dblExpenseToSalesYOYGrowth(i) = CalculateYOYGrowth(dblExpenseToSales(i), dblExpenseToSales(i + 1))
        strExpenseToSalesYOYGrowth(i) = Format(dblExpenseToSalesYOYGrowth(i), "0.0%")
        
        dblTaxRateYOYGrowth(i) = CalculateYOYGrowth(dblTaxRate(i), dblTaxRate(i + 1))
        strTaxRateYOYGrowth(i) = Format(dblTaxRateYOYGrowth(i), "0.0%")
        
        dblSharesYOYGrowth(i) = CalculateYOYGrowth(dblShares(i), dblShares(i + 1))
        strSharesYOYGrowth(i) = Format(dblSharesYOYGrowth(i), "0.0%")
    Next i
    
    With Range("Earnings")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="EPS = Net Income / Shares Outstanding" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Net Income              " & dblNetIncome(0) & "     " & dblNetIncome(1) & "     " & dblNetIncome(2) & "     " & dblNetIncome(3) & Chr(10) & _
                "YOY Net Income Growth   " & strNetIncomeYOYGrowth(0) & "     " & strNetIncomeYOYGrowth(1) & "     " & strNetIncomeYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Expense/Sales              " & strExpenseToSales(0) & "     " & strExpenseToSales(1) & "     " & strExpenseToSales(2) & "     " & strExpenseToSales(3) & Chr(10) & _
                "YOY Expense/Sales Growth   " & strExpenseToSalesYOYGrowth(0) & "     " & strExpenseToSalesYOYGrowth(1) & "     " & strExpenseToSalesYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Tax Rate              " & strTaxRate(0) & "     " & strTaxRate(1) & "     " & strTaxRate(2) & "     " & strTaxRate(3) & Chr(10) & _
                "YOY Tax Rate Growth   " & strTaxRateYOYGrowth(0) & "     " & strTaxRateYOYGrowth(1) & "     " & strTaxRateYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Shares Outstanding              " & dblShares(0) & "     " & dblShares(1) & "     " & dblShares(2) & "     " & dblShares(3) & Chr(10) & _
                "YOY Shares Outstanding Growth   " & strSharesYOYGrowth(0) & "     " & strSharesYOYGrowth(1) & "     " & strSharesYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With

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
    Dim i As Integer
    
    Range("EarningsYOYGrowth") = "YOY Growth (%)"
    
    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To (iYearsAvailableIncome - 2)
        dblYOYGrowth(i) = CalculateYOYGrowth(vEPS(i), vEPS(i + 1))
    Next i
    
    Call EvaluateEPSYOYGrowth(Range("EarningsYOYGrowth"), dblYOYGrowth)
    
End Sub

'===============================================================
' Procedure:    EvaluateEPSYOYGrowth
'
' Description:  Display YOY growth information.
'               if EPS < 0 or decreasing or growth is < EPS_GROWTH_MIN -> fail
'               else if EPS growth >= EPS_GROWTH_MIN -> pass
'               Scoring:
'               most recent year > EPS_GROWTH_MIN
'                   add EARNINGS_SCORE_MAX
'                   subtract if growth < 0
'               most recent year - 1 > EPS_GROWTH_MIN
'                   add EARNINGS_SCORE_MAX
'                   subtract if growth < 0
'               total score = score * EARNINGS_SCORE_WEIGHT
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   YOYGrowth As Range -> first cell of revenue YOY growth
'               YOY array -> YOY growth values
'                            YOY(0) is most recent year
'
' Returns:      N/A
'
' Rev History:  17Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateEPSYOYGrowth(YOYGrowth As Range, YOY() As Double)
    
    Dim i As Integer
    
    For i = 0 To (iYearsAvailableIncome - 2)
    YOYGrowth.Offset(0, i + 1).Select
    If vEPS(i) < 0 Or YOY(i) < 0 Or YOY(i) < EPS_GROWTH_MIN Then   'if EPS is negative or decreases or less than required
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultEarnings = FAIL
        If YOY(i) < 0 Then
            ScoreEarnings = ScoreEarnings - (EARNINGS_SCORE_MAX - i)
        End If
    ElseIf YOY(i + 1) - YOY(i) > 0.15 And i < (iYearsAvailableIncome - 1) Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultEarnings = FAIL
        ScoreEarnings = ScoreEarnings - (EARNINGS_SCORE_MAX - i)
    Else                                                            'if EPS growth is greater than required
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
        ScoreEarnings = ScoreEarnings + (EARNINGS_SCORE_MAX - i)
    End If
    YOYGrowth.Offset(0, i + 1) = YOY(i)
    Next i
    
    Range("I7").FormulaR1C1 = "=STDEV.P(RC[-6]:RC[-4])"
    If Range("I7").Value > 0.2 Then
        ScoreEarnings = ScoreEarnings - 10
    End If
    If ScoreEarnings < 0 Then
        ScoreEarnings = 0
    End If
    
    ScoreEarnings = ScoreEarnings * EARNINGS_SCORE_WEIGHT
    CheckEarningsPassFail
    EarningsScore

End Function

'===============================================================
' Procedure:    CheckEarningsPassFail
'
' Description:  Display check or x mark if the earnings
'               pass or fail the criteria
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  27Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CheckEarningsPassFail()

    If ResultEarnings = PASS Then
        Range("EarningsCheck") = CHECK_MARK
        Range("EarningsCheck").Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("EarningsCheck") = X_MARK
        Range("EarningsCheck").Font.ColorIndex = FONT_COLOR_RED
    End If

End Sub

'===============================================================
' Procedure:    EarningsScore
'
' Description:  Calculate score for earnings
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

Sub EarningsScore()

    Range("EarningsScore") = ScoreEarnings

End Sub

