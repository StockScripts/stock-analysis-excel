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
'               if ROE is greater than the required value -> green font
'               if ROE is positive but less than the required -> orange font
'               else ROE is negative -> red font
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
Sub EvaluateROE()

    Dim ErrorNum As Years   'used to catch errors for each year of data
    
    ResultGrowth = PASS
    
    DisplayROEInfo
        
    On Error GoTo ErrorHandler
    
    'populate ROE information
    ErrorNum = Year0
    
    'ROE = net income / equity
    dblROE(0) = dblNetIncome(0) / dblEquity(0)
    If dblROE(0) >= ROE_MIN Then    'if ROE is greater than required value
        Range("ROE").Offset(0, 1).Font.ColorIndex = FONT_COLOR_GREEN
    Else                            'if ROE is 0 or negative
        Range("ROE").Offset(0, 1).Font.ColorIndex = FONT_COLOR_RED
        ResultGrowth = FAIL
    End If
    Range("ROE").Offset(0, 1) = dblROE(0)
    
    ErrorNum = Year1
    
    dblROE(1) = dblNetIncome(1) / dblEquity(1)
    If dblROE(1) >= ROE_MIN Then
        Range("ROE").Offset(0, 2).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("ROE").Offset(0, 2).Font.ColorIndex = FONT_COLOR_RED
        ResultGrowth = FAIL
    End If
    Range("ROE").Offset(0, 2) = dblROE(1)
    
    ErrorNum = Year2
    
    dblROE(2) = dblNetIncome(2) / dblEquity(2)
    If dblROE(2) >= ROE_MIN Then
        Range("ROE").Offset(0, 3).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("ROE").Offset(0, 3).Font.ColorIndex = FONT_COLOR_RED
        ResultGrowth = FAIL
    End If
    Range("ROE").Offset(0, 3) = dblROE(2)
    
    ErrorNum = Year3
    
    dblROE(3) = dblNetIncome(3) / dblEquity(3)
    If dblROE(3) >= ROE_MIN Then
        Range("ROE").Offset(0, 4).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("ROE").Offset(0, 4).Font.ColorIndex = FONT_COLOR_RED
        ResultGrowth = FAIL
    End If
    Range("ROE").Offset(0, 4) = dblROE(3)
    
    CalculateROEYOYGrowth
    
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case Year0
            dblROE(0) = 0
            Range("ROE").Offset(0, 1) = dblROE(0)
        Case Year1
            dblROE(1) = 0
            Range("ROE").Offset(0, 2) = dblROE(1)
        Case Year2
            dblROE(2) = 0
            Range("ROE").Offset(0, 3) = dblROE(2)
        Case Year3
            dblROE(3) = 0
            Range("ROE").Offset(0, 4) = dblROE(3)
        Case Year4
            dblROE(4) = 0
            Range("ROE").Offset(0, 5) = dblROE(4)
   End Select
   
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
    
    Dim dblAssetTurnover(0 To 3) As Double
    Dim strAssetTurnover(0 To 3) As String
    Dim dblAssetTurnoverYOYGrowth(0 To 2) As Double
    Dim strAssetTurnoverYOYGrowth(0 To 2) As String
    Dim dblEquityYOYGrowth(0 To 2) As Double
    Dim strEquityYOYGrowth(0 To 2) As String
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
    
    'calculate Asset Turnover and YOY growth
    For i = 0 To 3
        'asset turnover = revenue/total assets
        dblAssetTurnover(i) = dblRevenue(i) / dblAssets(i)
        
        'if divide by 0
        If Err = ERROR_CODE_OVERFLOW Then
            dblAssetTurnover(i) = 0
        End If
        
        'convert to string to format and dispaly in comment box
        strAssetTurnover(i) = Format(dblAssetTurnover(i), "0.00")
    Next i
    
    For i = 0 To 2
        'calculated YOY growth
        dblAssetTurnoverYOYGrowth(i) = CalculateYOYGrowth(dblAssetTurnover(i), dblAssetTurnover(i + 1))
        
        'convert to string to format display in comment box
        strAssetTurnoverYOYGrowth(i) = Format(dblAssetTurnoverYOYGrowth(i), "0.0%")
    Next i
    
    'calculate equity YOY growth
    For i = 0 To 2
        dblEquityYOYGrowth(i) = CalculateYOYGrowth(dblEquity(i), dblEquity(i + 1))
        
        'convert to string to format display in comment box
        strEquityYOYGrowth(i) = Format(dblEquityYOYGrowth(i), "0.0%")
    Next i
    
    With Range("ROE")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="ROE = Net Income / Shareholder's Equity" & Chr(10) & _
                "Profit Margin x Assset Turnover x Leverage = ROE" & Chr(10) & _
                "Net Income/Sales x Sales/Assets x Assets/Equity = Net Income/Equity = ROE" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Asset Turnover              " & strAssetTurnover(0) & "     " & strAssetTurnover(1) & "     " & strAssetTurnover(2) & "     " & strAssetTurnover(3) & Chr(10) & _
                "YOY Asset Turnover Growth   " & strAssetTurnoverYOYGrowth(0) & "     " & strAssetTurnoverYOYGrowth(1) & "     " & strAssetTurnoverYOYGrowth(2) & "" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Equity              " & dblEquity(0) & "     " & dblEquity(1) & "     " & dblEquity(2) & "     " & dblEquity(3) & Chr(10) & _
                "YOY Equity Growth   " & strEquityYOYGrowth(0) & "     " & strEquityYOYGrowth(1) & "     " & strEquityYOYGrowth(2) & ""
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
