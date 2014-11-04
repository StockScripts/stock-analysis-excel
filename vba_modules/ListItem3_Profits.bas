Attribute VB_Name = "ListItem3_Profits"
Option Explicit

Private dblNetMargin(0 To 4) As Double
Private ResultProfits As Result

'===============================================================
' Procedure:    EvaluateNetMargin
'
' Description:  Display net margin information.
'               Call procedure to display YOY growth information
'               flag pass/fail for three most recent years
'               if net margin is negative -> red font -> fail
'               else -> green font -> pass
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
    
    Dim ErrorNum As Years   'used to catch errors for each year of data
    Dim i As Integer
    
    On Error Resume Next
    
    ResultProfits = PASS
    
    'net margin = net income / revenue
    For i = 0 To (iYearsAvailableIncome - 1)
        dblNetMargin(i) = dblNetIncome(i) / dblRevenue(i)
        If dblNetMargin(i) > 0 Then     'if net margin is positive
            Range("NetMargin").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_GREEN
        Else                            'if net margin is 0 or negative
            Range("NetMargin").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_RED
            ResultProfits = FAIL
        End If
        
        If Err = ERROR_CODE_OVERFLOW Then
            dblNetMargin(i) = 0
            Err.Clear
        End If
        Range("NetMargin").Offset(0, i + 1) = dblNetMargin(i)
    Next i
    
    DisplayProfitsInfo
    CalculateNetMarginYOYGrowth
    
End Sub

'===============================================================
' Procedure:    DisplayProfitsInfo
'
' Description:  Comment box information for Profits
'               - profits requirements
'               - net income and YOY growth
'               - net margin information
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  29Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub DisplayProfitsInfo()

    Dim dblNetIncomeYOYGrowth(0 To 2) As Double
    Dim strNetIncomeYOYGrowth(0 To 2) As String
    
    Dim dblTaxRate(0 To 3) As Double
    Dim strTaxRate(0 To 3) As String
    Dim dblTaxRateYOYGrowth(0 To 3) As Double
    Dim strTaxRateYOYGrowth(0 To 3) As String
    
    Dim dblExpenseToSales(0 To 3) As Double
    Dim strExpenseToSales(0 To 3) As String
    Dim dblExpenseToSalesYOYGrowth(0 To 3) As Double
    Dim strExpenseToSalesYOYGrowth(0 To 3) As String
    
    Dim i As Integer
    
    Range("ListItemNetMargin") = "Are profits increasing?"
    Range("NetMargin") = "Net Margin"
        
    With Range("ListItemNetMargin")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="What is it:" & Chr(10) & _
                "   Net profit margin is the percentage of revenue remaining after expenses have been deducted." & Chr(10) & _
                "Why is it important:" & Chr(10) & _
                "   Net margin measures how good a company is at converting revenue into profits." & Chr(10) & _
                "What to look for:" & Chr(10) & _
                "   Net margin should be stable or increasing." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   If net margin is increasing significantly, it could be due to a decrease in" & Chr(10) & _
                "   expenses or tax rate. Constantly cutting costs may not be sustainable in the long term."
        .Comment.Shape.TextFrame.AutoSize = True
    End With

    For i = 0 To (iYearsAvailableIncome - 1)
        dblExpenseToSales(i) = dblOperatingExpense(i) / dblRevenue(i)
        strExpenseToSales(i) = Format(dblExpenseToSales(i), "0.00")
        
        dblTaxRate(i) = 1 - (dblIncomeAfterTax(i) / dblIncomeBeforeTax(i))
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
    Next i
        
    With Range("NetMargin")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="Net Profit Margin = Net Income / Revenue" & Chr(10) & _
                "YOY Net Income              " & dblNetIncome(0) & "     " & dblNetIncome(1) & "     " & dblNetIncome(2) & "     " & dblNetIncome(3) & Chr(10) & _
                "YOY Net Income Growth   " & strNetIncomeYOYGrowth(0) & "     " & strNetIncomeYOYGrowth(1) & "     " & strNetIncomeYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Expense/Sales             " & strExpenseToSales(0) & "     " & strExpenseToSales(1) & "     " & strExpenseToSales(2) & "     " & strExpenseToSales(3) & Chr(10) & _
                "YOY Expense/Sales Growth   " & strExpenseToSalesYOYGrowth(0) & "     " & strExpenseToSalesYOYGrowth(1) & "     " & strExpenseToSalesYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Tax Rate             " & strTaxRate(0) & "     " & strTaxRate(1) & "     " & strTaxRate(2) & "     " & strTaxRate(3) & Chr(10) & _
                "YOY Tax Rate Growth   " & strTaxRateYOYGrowth(0) & "     " & strTaxRateYOYGrowth(1) & "     " & strTaxRateYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With
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

    Dim dblYOYGrowth(0 To 2) As Double
    Dim i As Integer
    
    Range("NetMarginYOYGrowth") = "YOY Growth (%)"
    
    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To (iYearsAvailableIncome - 2)
        dblYOYGrowth(i) = CalculateYOYGrowth(dblNetMargin(i), dblNetMargin(i + 1))
    Next i
    
    Call EvaluateNetMarginYOYGrowth(Range("NetMarginYOYGrowth"), dblYOYGrowth)
    
End Sub

'===============================================================
' Procedure:    EvaluateNetMarginYOYGrowth
'
' Description:  Display YOY growth information.
'               flag pass/fail for three most recent years
'               if net margin decreases -> red font -> fail
'               else net margin growth is positive -> green font -> pass
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
' Rev History:  17Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateNetMarginYOYGrowth(YOYGrowth As Range, YOY() As Double)
    
    Dim i As Integer
    
    For i = 0 To (iYearsAvailableIncome - 2)
        YOYGrowth.Offset(0, i + 1).Select
        If dblNetMargin(i) < 0 Or YOY(i) < 0 Then     'if net margin is negative or net margin decreases
            Selection.Font.ColorIndex = FONT_COLOR_RED
            ResultProfits = FAIL
        Else                                        'net margin is stable or increasing
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
        End If
        YOYGrowth.Offset(0, i + 1) = YOY(i)
    Next i
    
    CheckProfitsPassFail
    
End Function

'===============================================================
' Procedure:    CheckProfitsPassFail
'
' Description:  Display check or x mark if the profits
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
' Rev History:  29Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CheckProfitsPassFail()

    If ResultProfits = PASS Then
        Range("ProfitsCheck") = CHECK_MARK
        Range("ProfitsCheck").Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("ProfitsCheck") = X_MARK
        Range("ProfitsCheck").Font.ColorIndex = FONT_COLOR_RED
    End If

End Sub
