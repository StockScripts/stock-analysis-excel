Attribute VB_Name = "ListItem5_Equity"
Option Explicit

Private dblROE(0 To 4) As Double
Private ResultGrowth As Result
Private Const ROE_MIN = 0.1

'===============================================================
' Procedure:    EvaluateROE
'
' Description:  Display ROE information.
'               Call procedure to display YOY growth information
'               if recent year ROE is greater than the required value -> pass
'               else -> fail
'               if past year ROE is less than required -> warning
'               else -> pass
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
Sub EvaluateROE()
    
    Dim i As Integer
    
    ResultGrowth = PASS
    
    DisplayROEInfo
        
    On Error Resume Next
    
    'populate ROE information
    'ROE = net income / equity
    Range("ROE").Offset(0, 1).Select
    dblROE(0) = dblNetIncome(0) / dblEquity(0)
    If Err Then
        Selection.HorizontalAlignment = xlCenter
        Selection.Value = STR_NO_DATA
        Err.Clear
    Else
         If dblROE(0) >= ROE_MIN Then
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
        Else
            Selection.Font.ColorIndex = FONT_COLOR_RED
            ResultGrowth = FAIL
        End If
        Selection.Value = dblROE(0)
    End If
    
    For i = 1 To (iYearsAvailableIncome - 1)
        Range("ROE").Offset(0, i + 1).Select
        dblROE(i) = dblNetIncome(i) / dblEquity(i)
        If Err Then
            Selection.HorizontalAlignment = xlCenter
            Selection.Value = STR_NO_DATA
            Err.Clear
        Else
             If dblROE(i) >= ROE_MIN Then
                Selection.Font.ColorIndex = FONT_COLOR_GREEN
            Else
                Selection.Font.ColorIndex = FONT_COLOR_ORANGE
                ResultGrowth = FAIL
            End If
            Selection.Value = dblROE(i)
        End If
    Next i
    
    CalculateROEYOYGrowth

End Sub

'===============================================================
' Procedure:    DisplayROEInfo
'
' Description:  Comment box information for ROE
'               - ROE requirements
'               - equity and YOY growth
'               - dupont analysis
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  10Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub DisplayROEInfo()

    Dim dblEquityYOYGrowth(0 To 2) As Double
    Dim strEquityYOYGrowth(0 To 2) As String
    
    Dim dblNetIncomeYOYGrowth(0 To 2) As Double
    Dim strNetIncomeYOYGrowth(0 To 2) As String
    
    Dim dblProfitMargin(0 To 3) As Double
    Dim strProfitMargin(0 To 3) As String
    
    Dim dblProfitMarginYOYGrowth(0 To 2) As Double
    Dim strProfitMarginYOYGrowth(0 To 2) As String
    
    Dim dblAssetTurnover(0 To 3) As Double
    Dim strAssetTurnover(0 To 3) As String
    
    Dim dblAssetTurnoverYOYGrowth(0 To 2) As Double
    Dim strAssetTurnoverYOYGrowth(0 To 2) As String
    
    Dim dblLeverage(0 To 3) As Double
    Dim strLeverage(0 To 3) As String
    
    Dim dblLeverageYOYGrowth(0 To 2) As Double
    Dim strLeverageYOYGrowth(0 To 2) As String
    
    Dim i As Integer
    
    On Error Resume Next
    
    Range("ListItemROE") = "Is management effective?"
    Range("ROE") = "ROE"
    
    With Range("ListItemROE")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="What is it:" & Chr(10) & _
                "   Return on Equity is the net income as a percentage of shareholders equity." & Chr(10) & _
                "   It indicates how much the shareholders get for their investment in the company." & Chr(10) & _
                "Why is it important:" & Chr(10) & _
                "   Companies with high ROE and little debt are able to raise money for growth. " & Chr(10) & _
                "   It means they are able to invest back into the business without needing more capital." & Chr(10) & _
                "What to look for:" & Chr(10) & _
                "   ROE should be at least 10% and should not be decreasing." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   ROE can consist of three parts: profit margin, asset turnover, and leverage." & Chr(10) & _
                "   If ROE increases, make sure it is not increasing because the company is" & Chr(10) & _
                "   acquiring more debt and increasing leverage. If liabilities increase, equity" & Chr(10) & _
                "   decreases, which boosts ROE."
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    'calculate profit margin, Asset Turnover, and leverage
    For i = 0 To (iYearsAvailableIncome - 1)
        dblProfitMargin(i) = dblNetIncome(i) / dblRevenue(i)
        If Err Then
            dblProfitMargin(i) = 0
            Err.Clear
        End If
        'convert to string to format and display in comment box
        strProfitMargin(i) = Format(dblProfitMargin(i), "0.0%")
    
        dblAssetTurnover(i) = dblRevenue(i) / dblAssets(i)
        If Err Then
            dblAssetTurnover(i) = 0
            Err.Clear
        End If
        'convert to string to format and display in comment box
        strAssetTurnover(i) = Format(dblAssetTurnover(i), "0.00")
        
        dblLeverage(i) = dblAssets(i) / dblEquity(i)
        If Err Then
            dblLeverage(i) = 0
            Err.Clear
        End If
        'convert to string to format and display in comment box
        strLeverage(i) = Format(dblLeverage(i), "0.00")
    Next i
    
    'calculated YOY growth
    For i = 0 To (iYearsAvailableIncome - 2)
        dblNetIncomeYOYGrowth(i) = CalculateYOYGrowth(dblNetIncome(i), dblNetIncome(i + 1))
        'convert to string to format display in comment box
        strNetIncomeYOYGrowth(i) = Format(dblNetIncomeYOYGrowth(i), "0.0%")
        
        dblEquityYOYGrowth(i) = CalculateYOYGrowth(dblEquity(i), dblEquity(i + 1))
        'convert to string to format display in comment box
        strEquityYOYGrowth(i) = Format(dblEquityYOYGrowth(i), "0.0%")
    
        dblAssetTurnoverYOYGrowth(i) = CalculateYOYGrowth(dblAssetTurnover(i), dblAssetTurnover(i + 1))
        'convert to string to format display in comment box
        strAssetTurnoverYOYGrowth(i) = Format(dblAssetTurnoverYOYGrowth(i), "0.0%")
        
        dblProfitMarginYOYGrowth(i) = CalculateYOYGrowth(dblProfitMargin(i), dblProfitMargin(i + 1))
        'convert to string to format display in comment box
        strProfitMarginYOYGrowth(i) = Format(dblProfitMarginYOYGrowth(i), "0.0%")
        
        dblLeverageYOYGrowth(i) = CalculateYOYGrowth(dblLeverage(i), dblLeverage(i + 1))
        'convert to string to format display in comment box
        strLeverageYOYGrowth(i) = Format(dblLeverageYOYGrowth(i), "0.0%")
    Next i
    
    With Range("ROE")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="ROE = Net Income / Shareholder's Equity" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Net Income              " & dblNetIncome(0) & "     " & dblNetIncome(1) & "     " & dblNetIncome(2) & "     " & dblNetIncome(3) & Chr(10) & _
                "YOY Net Income Growth   " & strNetIncomeYOYGrowth(0) & "     " & strNetIncomeYOYGrowth(1) & "     " & strNetIncomeYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Equity              " & dblEquity(0) & "     " & dblEquity(1) & "     " & dblEquity(2) & "     " & dblEquity(3) & Chr(10) & _
                "YOY Equity Growth   " & strEquityYOYGrowth(0) & "     " & strEquityYOYGrowth(1) & "     " & strEquityYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "ROE = Net Income/Sales x Sales/Assets x Assets/Equity" & Chr(10) & _
                "       = Profit Margin x Assset Turnover x Leverage" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Profit Margin              " & strProfitMargin(0) & "     " & strProfitMargin(1) & "     " & strProfitMargin(2) & "     " & strProfitMargin(3) & Chr(10) & _
                "YOY Profit Margin Growth   " & strProfitMarginYOYGrowth(0) & "     " & strProfitMarginYOYGrowth(1) & "     " & strProfitMarginYOYGrowth(2) & "" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Asset Turnover              " & strAssetTurnover(0) & "       " & strAssetTurnover(1) & "       " & strAssetTurnover(2) & "       " & strAssetTurnover(3) & Chr(10) & _
                "YOY Asset Turnover Growth   " & strAssetTurnoverYOYGrowth(0) & "     " & strAssetTurnoverYOYGrowth(1) & "     " & strAssetTurnoverYOYGrowth(2) & "" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Leverage              " & strLeverage(0) & "       " & strLeverage(1) & "       " & strLeverage(2) & "       " & strLeverage(3) & Chr(10) & _
                "YOY Leverage Growth   " & strLeverageYOYGrowth(0) & "     " & strLeverageYOYGrowth(1) & "     " & strLeverageYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With

End Sub

'===============================================================
' Procedure:    CalculateROEYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for ROE data. Format cells.
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
Sub CalculateROEYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double
    Dim i As Integer
        
    Range("ROEYOYGrowth") = "YOY Growth (%)"
    
    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To 2
        dblYOYGrowth(i) = CalculateYOYGrowth(dblROE(i), dblROE(i + 1))
    Next i
    
    Call EvaluateROEYOYGrowth(Range("ROEYOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2))
    
End Sub

'===============================================================
' Procedure:    EvaluateROEYOYGrowth
'
' Description:  Display YOY growth information.
'               if ROE is less than required value and decreased from previous year -> red font
'               else if ROE decreased from previous year -> orange font
'               else ROE is stable or increased from previous year -> green font
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
' Rev History:  18Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateROEYOYGrowth(YOYGrowth As Range, YOY1, YOY2, YOY3)
    
    YOYGrowth.Offset(0, 3).Select
    If dblROE(2) < ROE_MIN And YOY3 < 0 Then            'if ROE is less than required and decreasing
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultGrowth = FAIL
    ElseIf YOY3 < 0 Then                                'if ROE is decreasing
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else                                                'ROE is stable or increasing
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If dblROE(1) < ROE_MIN And YOY2 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultGrowth = FAIL
    ElseIf YOY2 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If dblROE(0) < ROE_MIN And YOY1 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultGrowth = FAIL
    ElseIf YOY1 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
    CheckGrowthPassFail
    
End Function

'===============================================================
' Procedure:    CheckGrowthPassFail
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
Sub CheckGrowthPassFail()

    If ResultGrowth = PASS Then
        Range("ROECheck") = CHECK_MARK
        Range("ROECheck").Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("ROECheck") = X_MARK
        Range("ROECheck").Font.ColorIndex = FONT_COLOR_RED
    End If
    
End Sub
