Attribute VB_Name = "ListItem8_RedFlags"
Private Const RED_FLAG_GROWTH_MAX = 0.5
Private Const RECEIVABLES_MAX = 0.2
Private Const INVENTORY_MAX = 0.25

Private dblNetIncomeToOpCash(0 To 4) As Double
Private dblReceivablesToSales(0 To 4) As Double
Private dblInventoryToSales(0 To 4) As Double
Private dblSGAToSales(0 To 4) As Double
Private ResultRedFlags As Result
Private Const RED_FLAGS_SCORE_MAX = 4
Private Const RED_FLAGS_SCORE_WEIGHT = 1
Private ScoreRedFlags As Integer

'===============================================================
' Procedure:    EvaluateRedFlags
'
' Description:  Call procedures to evaluate red flag parameters
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateRedFlags()

    ResultRedFlags = PASS
    ScoreRedFlags = 0
    
    DisplayRedFlagsInfo
    
    EvaluateReceivablesToSales
    EvaluateInventoryToSales
    EvaluateSGAToSales
    EvaluateDividendPerShare
    
    CheckRedFlagsPassFail
    RedFlagsScore

End Sub

'===============================================================
' Procedure:    DisplayRedFlagsInfo
'
' Description:  Comment box information for red flags
'               - red flags information
'               - red flags requirements
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  31Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub DisplayRedFlagsInfo()
   
    Dim dblReceivablesYOYGrowth(0 To 3) As Double
    Dim strReceivablesYOYGrowth(0 To 3) As String
    
    Dim dblInventoryYOYGrowth(0 To 3) As Double
    Dim strInventoryYOYGrowth(0 To 3) As String
    
    Dim dblSGAYOYGrowth(0 To 3) As Double
    Dim strSGAYOYGrowth(0 To 3) As String
    
    Dim dblRevenueYOYGrowth(0 To 3) As Double
    Dim strRevenueYOYGrowth(0 To 3) As String
    
    Range("ListItemRedFlags") = "Are there red flags?"
    
    Range("Receivables") = "Receivables/Sales"
    Range("ReceivablesYOYGrowth") = "YOY Growth (%)"
    
    Range("Inventory") = "Inventory/Sales"
    Range("InventoryYOYGrowth") = "YOY Growth (%)"
    
    Range("SGA") = "SGA/Sales"
    Range("SGAYOYGrowth") = "YOY Growth (%)"
    
    Range("Dividend") = "Dividend/Share"
    Range("DividendYOYGrowth") = "YOY Growth (%)"

    'red flags info
    With Range("ListItemRedFlags")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="What is it:" & Chr(10) & _
                "   A red flag is anything that marks a stock as undesirable and may differ depending" & Chr(10) & _
                "   on the type of company." & Chr(10) & _
                "Why is it important:" & Chr(10) & _
                "   A red flag indicates potential problems within a company." & Chr(10) & _
                "What to look for:" & Chr(10) & _
                "   Accounts receivable should not exceed 20% of annual sales." & Chr(10) & _
                "   Inventory should not exceed 25% of cost of goods sold." & Chr(10) & _
                "   Inventory, accounts receivables, and SGA should not grow faster than sales." & Chr(10) & _
                "   Dividend per share should not decrease." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   Inventory, sales, and receivables should move in tandem because customers" & Chr(10) & _
                "   do not pay up front if they can avoid it." & Chr(10) & _
                "   Inflating the inventory may increase earnings. Inventory fraud is a way to produce instant earnings." & Chr(10) & _
                "   Increasing SGA can be indicate operational problems along with deteriorating operating margins."
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    'calculate YOY growth
    For i = 0 To (iYearsAvailableIncome - 2)
        dblReceivablesYOYGrowth(i) = CalculateYOYGrowth(dblReceivables(i), dblReceivables(i + 1))
        strReceivablesYOYGrowth(i) = Format(dblReceivablesYOYGrowth(i), "0.0%")
        
        dblInventoryYOYGrowth(i) = CalculateYOYGrowth(dblInventory(i), dblInventory(i + 1))
        strInventoryYOYGrowth(i) = Format(dblInventoryYOYGrowth(i), "0.0%")
        
        dblSGAYOYGrowth(i) = CalculateYOYGrowth(dblSGA(i), dblSGA(i + 1))
        strSGAYOYGrowth(i) = Format(dblSGAYOYGrowth(i), "0.0%")
        
        dblRevenueYOYGrowth(i) = CalculateYOYGrowth(dblRevenue(i), dblRevenue(i + 1))
        strRevenueYOYGrowth(i) = Format(dblRevenueYOYGrowth(i), "0.0%")
    Next i
    
    With Range("Receivables")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="YOY Receivables" & "                " & dblReceivables(0) & "       " & dblReceivables(1) & "       " & dblReceivables(2) & "       " & dblReceivables(3) & Chr(10) & _
                "YOY Receivables Growth     " & strReceivablesYOYGrowth(0) & "     " & strReceivablesYOYGrowth(1) & "     " & strReceivablesYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Revenue              " & dblRevenue(0) & "     " & dblRevenue(1) & "     " & dblRevenue(2) & "     " & dblRevenue(3) & Chr(10) & _
                "YOY Revenue Growth   " & strRevenueYOYGrowth(0) & "     " & strRevenueYOYGrowth(1) & "     " & strRevenueYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    With Range("Inventory")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="YOY Inventory" & "                " & dblInventory(0) & "       " & dblInventory(1) & "       " & dblInventory(2) & "       " & dblInventory(3) & Chr(10) & _
                "YOY Inventory Growth     " & strInventoryYOYGrowth(0) & "     " & strInventoryYOYGrowth(1) & "     " & strInventoryYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Revenue              " & dblRevenue(0) & "     " & dblRevenue(1) & "     " & dblRevenue(2) & "     " & dblRevenue(3) & Chr(10) & _
                "YOY Revenue Growth   " & strRevenueYOYGrowth(0) & "     " & strRevenueYOYGrowth(1) & "     " & strRevenueYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    With Range("SGA")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="YOY SGA" & "                " & dblSGA(0) & "       " & dblSGA(1) & "       " & dblSGA(2) & "       " & dblSGA(3) & Chr(10) & _
                "YOY SGA Growth     " & strSGAYOYGrowth(0) & "     " & strSGAYOYGrowth(1) & "     " & strSGAYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Revenue              " & dblRevenue(0) & "     " & dblRevenue(1) & "     " & dblRevenue(2) & "     " & dblRevenue(3) & Chr(10) & _
                "YOY Revenue Growth   " & strRevenueYOYGrowth(0) & "     " & strRevenueYOYGrowth(1) & "     " & strRevenueYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
End Sub


'===============================================================
' Procedure:    EvaluateReceivablesToSales
'
' Description:  Display Receivables to Sales information.
'               if recent year receivables/sales > RECEIVABLES_MAX -> fail
'               else pass
'
'               if past years receivables/sales > RECEIVABLES_MAX -> warning
'
'               Call procedure to display YOY growth information
'
'               catch errors and set value to STR_NO_DATA
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateReceivablesToSales()

    Dim i As Integer
    
    On Error Resume Next
        
    dblReceivablesToSales(0) = dblReceivables(0) / dblRevenue(0)
    Range("Receivables").Offset(0, 1).Select
    If Err Then
        Selection.HorizontalAlignment = xlCenter
        Selection.Value = STR_NO_DATA
        Err.Clear
    Else
        If dblReceivablesToSales(0) > RECEIVABLES_MAX Then
            Selection.Font.ColorIndex = FONT_COLOR_RED
            ResultRedFlags = FAIL
        Else
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreRedFlags = ScoreRedFlags + (RED_FLAGS_SCORE_MAX - i)
        End If
        Selection.Value = dblReceivablesToSales(0)
    End If
        
    'populate Receivables/Sales information
    For i = 1 To (iYearsAvailableIncome - 1)
        dblReceivablesToSales(i) = dblReceivables(i) / dblRevenue(i)
        Range("Receivables").Offset(0, i + 1).Select
        If Err Then
            Selection.HorizontalAlignment = xlCenter
            Selection.Value = STR_NO_DATA
            Err.Clear
        Else
            If dblReceivablesToSales(i) > RECEIVABLES_MAX Then
                Selection.Font.ColorIndex = FONT_COLOR_ORANGE
            Else
                Selection.Font.ColorIndex = FONT_COLOR_GREEN
                ScoreRedFlags = ScoreRedFlags + (RED_FLAGS_SCORE_MAX - i)
            End If
            Selection.Value = dblReceivablesToSales(i)
        End If
    Next i

    CalculateReceivablesToSalesYOYGrowth

End Sub

'===============================================================
' Procedure:    CalculateReceivablesToSalesYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for Receivables to Sales data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateReceivablesToSalesYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double

    On Error Resume Next

    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To (iYearsAvailableIncome - 2)
        dblYOYGrowth(i) = CalculateYOYGrowth(dblReceivablesToSales(i), dblReceivablesToSales(i + 1))
        
        If Err Then
            dblYOYGrowth(i) = 0
            Err.Clear
        End If
    Next i

    Call EvaluateRedFlagYOYGrowth(Range("ReceivablesYOYGrowth"), dblYOYGrowth)
    
End Sub

'===============================================================
' Procedure:    EvaluateInventoryToSales
'
' Description:  Display Inventory to Sales information.
'               if recent year inventory/sales > INVENTORY_MAX -> fail
'               else pass
'
'               if past years inventory/sales > INVENTORY_MAX -> warning
'
'               Call procedure to display YOY growth information
'
'               catch errors and set value to STR_NO_DATA
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateInventoryToSales()

    Dim i As Integer
    
    On Error Resume Next
    
    dblInventoryToSales(0) = dblInventory(0) / dblRevenue(0)
    Range("Inventory").Offset(0, 1).Select
    If Err Then
        Selection.HorizontalAlignment = xlCenter
        Selection.Value = STR_NO_DATA
        Err.Clear
    Else
        If dblInventoryToSales(0) > INVENTORY_MAX Then
            Selection.Font.ColorIndex = FONT_COLOR_RED
        Else
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreRedFlags = ScoreRedFlags + (RED_FLAGS_SCORE_MAX - i)
        End If
        Selection.Value = dblInventoryToSales(0)
    End If
        
    'populate Inventory/Sales information
    For i = 1 To (iYearsAvailableIncome - 1)
        dblInventoryToSales(i) = dblInventory(i) / dblRevenue(i)
        Range("Inventory").Offset(0, i + 1).Select
        If Err Then
            Selection.HorizontalAlignment = xlCenter
            Selection.Value = STR_NO_DATA
            Err.Clear
        Else
            If dblInventoryToSales(i) > INVENTORY_MAX Then
                Selection.Font.ColorIndex = FONT_COLOR_ORANGE
            Else
                Selection.Font.ColorIndex = FONT_COLOR_GREEN
                ScoreRedFlags = ScoreRedFlags + (RED_FLAGS_SCORE_MAX - i)
            End If
            Selection.Value = dblInventoryToSales(i)
        End If
    Next i

    CalculateInventoryToSalesYOYGrowth

End Sub

'===============================================================
' Procedure:    CalculateInventoryToSalesYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for Inventory to Sales data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateInventoryToSalesYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double
    Dim i As Integer
    
    On Error Resume Next

    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To (iYearsAvailableIncome - 2)
        dblYOYGrowth(i) = CalculateYOYGrowth(dblInventoryToSales(i), dblInventoryToSales(i + 1))
        
        If Err Then
            dblYOYGrowth(i) = 0
            Err.Clear
        End If
    Next i
    
    Call EvaluateRedFlagYOYGrowth(Range("InventoryYOYGrowth"), dblYOYGrowth)
    
End Sub

'===============================================================
' Procedure:    EvaluateSGAToSales
'
' Description:  Display SGA to Sales information.
'               Call procedure to display YOY growth information
'               catch errors and set value to STR_NO_DATA
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateSGAToSales()

    Dim i As Integer
    
    On Error Resume Next
    
    'populate SGA to Sales information
    For i = 0 To (iYearsAvailableIncome - 1)
        dblSGAToSales(i) = dblSGA(i) / dblRevenue(i)
        Range("SGA").Offset(0, i + 1).Select
        If Err Then
            Selection.HorizontalAlignment = xlCenter
            Selection.Value = STR_NO_DATA
            Err.Clear
        Else
            Selection.Value = dblSGAToSales(i)
        End If
    Next i
    
    CalculateSGAToSalesYOYGrowth

End Sub

'===============================================================
' Procedure:    CalculateSGAToSalesYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for SGA to Sales data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateSGAToSalesYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double
    Dim i As Integer

    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To (iYearsAvailableIncome - 2)
        dblYOYGrowth(i) = CalculateYOYGrowth(dblSGAToSales(i), dblSGAToSales(i + 1))
        
        If Err Then
            dblYOYGrowth(i) = 0
            Err.Clear
        End If
    Next i
    
    Call EvaluateRedFlagYOYGrowth(Range("SGAYOYGrowth"), dblYOYGrowth)
    
End Sub

'===============================================================
' Procedure:    EvaluateRedFlagYOYGrowth
'
' Description:  Display YOY growth information.
'               if YOY growth is greater than max value -> red font
'               else YOY growth is decreasing -> green font
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
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateRedFlagYOYGrowth(YOYGrowth As Range, YOY() As Double)
    
    Dim i As Integer
    
    For i = 0 To (iYearsAvailableIncome - 2)
        YOYGrowth.Offset(0, i + 1).Select
        If YOY(i) > RED_FLAG_GROWTH_MAX Then
            Selection.Font.ColorIndex = FONT_COLOR_RED
            ResultRedFlags = FAIL
        Else
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreRedFlags = ScoreRedFlags + (RED_FLAGS_SCORE_MAX - i)
        End If
        Selection.Value = YOY(i)
    Next i

End Function

'===============================================================
' Procedure:    EvaluateDividendPerShare
'
' Description:  Display dividend per share information.
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
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateDividendPerShare()
    
    Dim i As Integer
    
    'populate Dividend Per Share information
    For i = 0 To (iYearsAvailableIncome - 1)
        Range("Dividend").Offset(0, i + 1) = dblDividendPerShare(i)
    Next i
    
    CalculateDividendPerShareYOYGrowth
    
End Sub

'===============================================================
' Procedure:    CalculateDividendPerShareYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for dividend data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateDividendPerShareYOYGrowth()

    Dim dblYOYGrowth(0 To 2) As Double
    Dim i As Integer

    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To (iYearsAvailableIncome - 2)
        dblYOYGrowth(i) = CalculateYOYGrowth(dblDividendPerShare(i), dblDividendPerShare(i + 1))
        
        If Err Then
            dblYOYGrowth(i) = 0
            Err.Clear
        End If
    Next i
    
    Call EvaluateDivPerShareYOYGrowth(Range("DividendYOYGrowth"), dblYOYGrowth)
    
    
End Sub

'===============================================================
' Procedure:    EvaluateDivPerShareYOYGrowth
'
' Description:  Display YOY growth information.
'               if most recent year dividend is decreasing -> red font
'               else dividend is increasing -> green font
'
'               if past year dividend is decreasing -> warning
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
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateDivPerShareYOYGrowth(YOYGrowth As Range, YOY() As Double)
    
    Dim i As Integer
    
    YOYGrowth.Offset(0, 1).Select
    If YOY(0) < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultRedFlags = FAIL
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
        ScoreRedFlags = ScoreRedFlags + (RED_FLAGS_SCORE_MAX - i)
    End If
    YOYGrowth.Offset(0, 1) = YOY(0)
    
    For i = 1 To (iYearsAvailableIncome - 2)
        YOYGrowth.Offset(0, i + 1).Select
        If YOY(i) < 0 Then
            Selection.Font.ColorIndex = FONT_COLOR_ORANGE
        Else
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreRedFlags = ScoreRedFlags + (RED_FLAGS_SCORE_MAX - i)
        End If
        YOYGrowth.Offset(0, i + 1) = YOY(i)
    Next i
    
End Function

'===============================================================
' Procedure:    CheckRedFlagsPassFail
'
' Description:  Display check or x mark if the red flags
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
' Rev History:  15Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CheckRedFlagsPassFail()

    If ResultRedFlags = PASS Then
        Range("RedFlagsCheck") = CHECK_MARK
        Range("RedFlagsCheck").Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("RedFlagsCheck") = X_MARK
        Range("RedFlagsCheck").Font.ColorIndex = FONT_COLOR_RED
    End If

End Sub

'===============================================================
' Procedure:    RedFlagsScore
'
' Description:  Calculate score for red flags
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

Sub RedFlagsScore()

    Range("RedFlagsScore") = ScoreRedFlags
End Sub





