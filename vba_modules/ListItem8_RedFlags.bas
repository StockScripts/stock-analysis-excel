Attribute VB_Name = "ListItem8_RedFlags"
Private Const RED_FLAG_GROWTH_MAX = 0.5
Private Const RECEIVABLES_MAX = 0.2
Private Const INVENTORY_MAX = 0.25

Private dblNetIncomeToOpCash(0 To 4) As Double
Private dblReceivablesToSales(0 To 4) As Double
Private dblInventoryToSales(0 To 4) As Double
Private dblSGAToSales(0 To 4) As Double
Private ResultRedFlags As Result

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
    
    DisplayRedFlagsInfo
    
    EvaluateNetIncomeToOpCash
    EvaluateReceivablesToSales
    EvaluateInventoryToSales
    EvaluateSGAToSales
    EvaluateDividendPerShare
    
    CheckRedFlagsPassFail

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

    Dim dblNetIncomeYOYGrowth(0 To 3) As Double
    Dim strNetIncomeYOYGrowth(0 To 3) As String
    
    Dim dblOpCashFlowYOYGrowth(0 To 3) As Double
    Dim strOpCashFlowYOYGrowth(0 To 3) As String
    
    Dim dblReceivablesYOYGrowth(0 To 3) As Double
    Dim strReceivablesYOYGrowth(0 To 3) As String
    
    Dim dblRevenueYOYGrowth(0 To 3) As Double
    Dim strRevenueYOYGrowth(0 To 3) As String
    
    Range("ListItemRedFlags") = "Are there red flags?"
    Range("NetIncomeToOpCash") = "Income/Op Cash"
    Range("Receivables") = "Receivables/Sales"
    Range("Inventory") = "Inventory/Sales"
    Range("SGA") = "SGA/Sales"
    Range("Dividend") = "Dividend/Share"

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
                "   Net income should not be increasing while cash flow from operations is decreasing." & Chr(10) & _
                "   Accounts receivable should not exceed 20% of annual sales." & Chr(10) & _
                "   Inventory should not exceed 25% of cost of goods sold." & Chr(10) & _
                "   Inventory, accounts receivables, and SGA should not grow faster than sales." & Chr(10) & _
                "   Dividend per share should not decrease." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   Inventory, sales, and receivables should move in tandem because customers" & Chr(10) & _
                "   do not pay up front if they can avoid it." & Chr(10) & _
                "   Inflating the inventory may increase earnings. Inventory fraud is a way to produce instant earnings."
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    'calculate YOY growth
    For i = 0 To (iYearsAvailableIncome - 2)
        dblNetIncomeYOYGrowth(i) = CalculateYOYGrowth(dblNetIncome(i), dblNetIncome(i + 1))
        strNetIncomeYOYGrowth(i) = Format(dblNetIncomeYOYGrowth(i), "0.0%")
        
        dblOpCashFlowYOYGrowth(i) = CalculateYOYGrowth(dblOpCashFlow(i), dblOpCashFlow(i + 1))
        strOpCashFlowYOYGrowth(i) = Format(dblOpCashFlowYOYGrowth(i), "0.0%")
        
        dblReceivablesYOYGrowth(i) = CalculateYOYGrowth(dblReceivables(i), dblReceivables(i + 1))
        strReceivablesYOYGrowth(i) = Format(dblReceivablesYOYGrowth(i), "0.0%")
        
        dblRevenueYOYGrowth(i) = CalculateYOYGrowth(dblRevenue(i), dblRevenue(i + 1))
        strRevenueYOYGrowth(i) = Format(dblRevenueYOYGrowth(i), "0.0%")
    Next i
    
    With Range("NetIncomeToOpCash")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="YOY Net Income                " & dblNetIncome(0) & "     " & dblNetIncome(1) & "     " & dblNetIncome(2) & "     " & dblNetIncome(3) & Chr(10) & _
                "YOY Net Income Growth     " & strNetIncomeYOYGrowth(0) & "     " & strNetIncomeYOYGrowth(1) & "     " & strNetIncomeYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Op Cash Flow              " & dblOpCashFlow(0) & "     " & dblOpCashFlow(1) & "     " & dblOpCashFlow(2) & "     " & dblOpCashFlow(3) & Chr(10) & _
                "YOY Op Cash Flow Growth   " & strOpCashFlowYOYGrowth(0) & "     " & strOpCashFlowYOYGrowth(1) & "     " & strOpCashFlowYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    With Range("Receivables")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="YOY Receivables                " & dblReceivables(0) & "     " & dblReceivables(1) & "     " & dblReceivables(2) & "     " & dblReceivables(3) & Chr(10) & _
                "YOY Receivables Growth     " & strReceivablesYOYGrowth(0) & "     " & strReceivablesYOYGrowth(1) & "     " & strReceivablesYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Revenue              " & dblRevenue(0) & "     " & dblRevenue(1) & "     " & dblRevenue(2) & "     " & dblRevenue(3) & Chr(10) & _
                "YOY Revenue Growth   " & strRevenueYOYGrowth(0) & "     " & strRevenueYOYGrowth(1) & "     " & strRevenueYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
End Sub

'===============================================================
' Procedure:    EvaluateNetIncomeToOpCash
'
' Description:  Display Net Income to Operating Cash flow information.
'               Call procedure to display YOY growth information
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
' Rev History:  01Nov14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateNetIncomeToOpCash()
       
'   populate Receivables/Sales information
    dblNetIncomeToOpCash(0) = dblNetIncome(0) / dblOpCashFlow(0)
    Range("NetIncometoOpcash").Offset(0, 1) = dblNetIncomeToOpCash(0)
    
    dblNetIncomeToOpCash(1) = dblNetIncome(1) / dblOpCashFlow(1)
    Range("NetIncometoOpcash").Offset(0, 2) = dblNetIncomeToOpCash(1)
    
    dblNetIncomeToOpCash(2) = dblNetIncome(2) / dblOpCashFlow(2)
    Range("NetIncometoOpcash").Offset(0, 3) = dblNetIncomeToOpCash(2)
    
    dblNetIncomeToOpCash(3) = dblNetIncome(3) / dblOpCashFlow(3)
    Range("NetIncometoOpcash").Offset(0, 4) = dblNetIncomeToOpCash(3)

    CalculateNetIncomeToOpCashYOYGrowth

End Sub

'===============================================================
' Procedure:    CalculateNetIncomeToOpCashYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for Net Income to operating cash flow data.
'               Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  01Nov14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateNetIncomeToOpCashYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double

    Range("NetIncomeToOpCashYOYGrowth") = "YOY Growth (%)"

    'populate YOY growth information
    '(0) is most recent year
    dblYOYGrowth(0) = CalculateYOYGrowth(dblNetIncomeToOpCash(0), dblNetIncomeToOpCash(1))
    dblYOYGrowth(1) = CalculateYOYGrowth(dblNetIncomeToOpCash(1), dblNetIncomeToOpCash(2))
    dblYOYGrowth(2) = CalculateYOYGrowth(dblNetIncomeToOpCash(2), dblNetIncomeToOpCash(3))
    
    Call EvaluateRedFlagYOYGrowth(Range("NetIncomeToOpCashYOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2))
    

End Sub

'===============================================================
' Procedure:    EvaluateReceivablesToSales
'
' Description:  Display Receivables to Sales information.
'               Call procedure to display YOY growth information
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
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateReceivablesToSales()

    Dim ErrorNum As Years   'used to catch errors for each year of data
    
    On Error GoTo ErrorHandler
        
'   populate Receivables/Sales information
    ErrorNum = Year0
    dblReceivablesToSales(0) = dblReceivables(0) / dblRevenue(0)
    Range("Receivables").Offset(0, 1) = dblReceivablesToSales(0)
    
    ErrorNum = Year1
    dblReceivablesToSales(1) = dblReceivables(1) / dblRevenue(1)
    Range("Receivables").Offset(0, 2) = dblReceivablesToSales(1)
    
    ErrorNum = Year2
    dblReceivablesToSales(2) = dblReceivables(2) / dblRevenue(2)
    Range("Receivables").Offset(0, 3) = dblReceivablesToSales(2)
    
    ErrorNum = Year3
    dblReceivablesToSales(3) = dblReceivables(3) / dblRevenue(3)
    Range("Receivables").Offset(0, 4) = dblReceivablesToSales(3)

    CalculateReceivablesToSalesYOYGrowth
    
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case Year0
            dblReceivablesToSales(0) = 0
            Range("Receivables").Offset(0, 1) = dblReceivablesToSales(0)
        Case Year1
            dblReceivablesToSales(1) = 0
            Range("Receivables").Offset(0, 2) = dblReceivablesToSales(1)
        Case Year2
            dblReceivablesToSales(2) = 0
            Range("Receivables").Offset(0, 3) = dblReceivablesToSales(2)
        Case Year3
            dblReceivablesToSales(3) = 0
            Range("Receivables").Offset(0, 4) = dblReceivablesToSales(3)
   End Select
   
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

    Range("ReceivablesYOYGrowth") = "YOY Growth (%)"

    'populate YOY growth information
    '(0) is most recent year
    dblYOYGrowth(0) = CalculateYOYGrowth(dblReceivablesToSales(0), dblReceivablesToSales(1))
    dblYOYGrowth(1) = CalculateYOYGrowth(dblReceivablesToSales(1), dblReceivablesToSales(2))
    dblYOYGrowth(2) = CalculateYOYGrowth(dblReceivablesToSales(2), dblReceivablesToSales(3))
    
    Call EvaluateRedFlagYOYGrowth(Range("ReceivablesYOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2))
    
End Sub

'===============================================================
' Procedure:    EvaluateInventoryToSales
'
' Description:  Display Inventory to Sales information.
'               Call procedure to display YOY growth information
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
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateInventoryToSales()

    Dim ErrorNum As Years   'used to catch errors for each year of data
    
    On Error GoTo ErrorHandler
    
'   populate ROE information
    ErrorNum = Year0
    dblInventoryToSales(0) = dblInventory(0) / dblRevenue(0)
    Range("Inventory").Offset(0, 1) = dblInventoryToSales(0)
    
    ErrorNum = Year1
    dblInventoryToSales(1) = dblInventory(1) / dblRevenue(1)
    Range("Inventory").Offset(0, 2) = dblInventoryToSales(1)
    
    ErrorNum = Year2
    dblInventoryToSales(2) = dblInventory(2) / dblRevenue(2)
    Range("Inventory").Offset(0, 3) = dblInventoryToSales(2)
    
    ErrorNum = Year3
    dblInventoryToSales(3) = dblInventory(3) / dblRevenue(3)
    Range("Inventory").Offset(0, 4) = dblInventoryToSales(3)

    CalculateInventoryToSalesYOYGrowth
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case Year0
            dblInventoryToSales(0) = 0
            Range("Inventory").Offset(0, 1) = dblInventoryToSales(0)
        Case Year1
            dblInventoryToSales(1) = 0
            Range("Inventory").Offset(0, 2) = dblInventoryToSales(1)
        Case Year2
            dblInventoryToSales(2) = 0
            Range("Inventory").Offset(0, 3) = dblInventoryToSales(2)
        Case Year3
            dblInventoryToSales(3) = 0
            Range("Inventory").Offset(0, 4) = dblInventoryToSales(3)
   End Select
   
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
    
    Range("InventoryYOYGrowth") = "YOY Growth (%)"

    'populate YOY growth information
    '(0) is most recent year
    dblYOYGrowth(0) = CalculateYOYGrowth(dblInventoryToSales(0), dblInventoryToSales(1))
    dblYOYGrowth(1) = CalculateYOYGrowth(dblInventoryToSales(1), dblInventoryToSales(2))
    dblYOYGrowth(2) = CalculateYOYGrowth(dblInventoryToSales(2), dblInventoryToSales(3))
    
    Call EvaluateRedFlagYOYGrowth(Range("InventoryYOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2))
    
End Sub

'===============================================================
' Procedure:    EvaluateSGAToSales
'
' Description:  Display SGA to Sales information.
'               Call procedure to display YOY growth information
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
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateSGAToSales()

    Dim ErrorNum As Years   'used to catch errors for each year of data
    
    On Error GoTo ErrorHandler
    
'   populate SGA to Sales information
    ErrorNum = Year0
    dblSGAToSales(0) = dblSGA(0) / dblRevenue(0)
    Range("SGA").Offset(0, 1) = dblSGAToSales(0)
    
    ErrorNum = Year1
    dblSGAToSales(1) = dblSGA(1) / dblRevenue(1)
    Range("SGA").Offset(0, 2) = dblSGAToSales(1)
    
    ErrorNum = Year2
    dblSGAToSales(2) = dblSGA(2) / dblRevenue(2)
    Range("SGA").Offset(0, 3) = dblSGAToSales(2)
    
    ErrorNum = Year3
    dblSGAToSales(3) = dblSGA(3) / dblRevenue(3)
    Range("SGA").Offset(0, 4) = dblSGAToSales(3)

    Range("SGA").AddComment
    Range("SGA").Comment.Visible = False
    Range("SGA").Comment.Text Text:="Overhead costs" & Chr(10) & _
                "operating expenses except cost of sales, " & Chr(10) & _
                "R&D, and depreciation and amortization." & Chr(10) & _
                "can be used to detect operational problems along with deteriorating operating margins" & Chr(10) & _
                "SGA/Sales should be stable and not increasing"
    Range("SGA").Comment.Shape.TextFrame.AutoSize = True
    
    CalculateSGAToSalesYOYGrowth
    
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case Year0
            dblSGAToSales(0) = 0
            Range("SGA").Offset(0, 1) = dblSGAToSales(0)
        Case Year1
            dblSGAToSales(1) = 0
            Range("SGA").Offset(0, 2) = dblSGAToSales(1)
        Case Year2
            dblSGAToSales(2) = 0
            Range("SGA").Offset(0, 3) = dblSGAToSales(2)
        Case Year3
            dblSGAToSales(3) = 0
            Range("SGA").Offset(0, 4) = dblSGAToSales(3)
   End Select
   
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
    
    Range("SGAYOYGrowth") = "YOY Growth (%)"

    'populate YOY growth information
    '(0) is most recent year
    dblYOYGrowth(0) = CalculateYOYGrowth(dblSGAToSales(0), dblSGAToSales(1))
    dblYOYGrowth(1) = CalculateYOYGrowth(dblSGAToSales(1), dblSGAToSales(2))
    dblYOYGrowth(2) = CalculateYOYGrowth(dblSGAToSales(2), dblSGAToSales(3))
    
    Call EvaluateRedFlagYOYGrowth(Range("SGAYOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2))
    
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
'               YOY1, YOY2, YOY3, YOY4 -> YOY growth values
'                                         (YOY1 is most recent year)
'
' Returns:      N/A
'
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateRedFlagYOYGrowth(YOYGrowth As Range, YOY1, YOY2, YOY3)
    
    YOYGrowth.Offset(0, 3).Select
    If YOY3 > RED_FLAG_GROWTH_MAX Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultRedFlags = FAIL
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
        If YOY2 > RED_FLAG_GROWTH_MAX Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultRedFlags = FAIL
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If YOY1 > RED_FLAG_GROWTH_MAX Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultRedFlags = FAIL
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
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
    
'   populate Dividend Per Share information
    Range("Dividend").Offset(0, 1) = dblDividendPerShare(0)
    Range("Dividend").Offset(0, 2) = dblDividendPerShare(1)
    Range("Dividend").Offset(0, 3) = dblDividendPerShare(2)
    Range("Dividend").Offset(0, 4) = dblDividendPerShare(3)
    
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
    
    Range("DividendYOYGrowth") = "YOY Growth (%)"

    'populate YOY growth information
    '(0) is most recent year
    dblYOYGrowth(0) = CalculateYOYGrowth(dblDividendPerShare(0), dblDividendPerShare(1))
    dblYOYGrowth(1) = CalculateYOYGrowth(dblDividendPerShare(1), dblDividendPerShare(2))
    dblYOYGrowth(2) = CalculateYOYGrowth(dblDividendPerShare(2), dblDividendPerShare(3))
    
    Call EvaluateDivPerShareYOYGrowth(Range("DividendYOYGrowth"), dblYOYGrowth(0), dblYOYGrowth(1), dblYOYGrowth(2))
    
    
End Sub

'===============================================================
' Procedure:    EvaluateDivPerShareYOYGrowth
'
' Description:  Display YOY growth information.
'               if dividend is decreasing -> red font
'               else dividend is increasing -> green font
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
' Rev History:  20Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateDivPerShareYOYGrowth(YOYGrowth As Range, YOY1, YOY2, YOY3)
    
    YOYGrowth.Offset(0, 3).Select
    If YOY3 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultRedFlags = FAIL
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If YOY2 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultRedFlags = FAIL
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If YOY1 < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultRedFlags = FAIL
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
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






