Attribute VB_Name = "Checklist"
Option Explicit

Private Const COLUMN_WIDTH_ANNUAL_DATA = 12
Private Const COLUMN_WIDTH_CHECKLIST_ITEM = 17
Private Const COLOR_LIGHT_BLUE = 24

'===============================================================
' Procedure:    CreateChecklistStockAnalysis
'
' Description:  Call procedures create stock analysis worksheet,
'               format worksheet, and generate checklist information.
'               - revenue
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   16Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CreateChecklistStockAnalysis()

    CreateWorkSheetStockAnalysis
    AssignCellItemsStockAnalysis
    FormatCheckListStockAnalysis
      
    EvaluateRevenue
    EvaluateEPS
    EvaluateNetMargin
    EvaluateFreeCashFlow
    EvaluateROE
    EvaluateFinancialLeverage
    EvaluateQuickRatio
    EvaluateRedFlags
    EvaluatePrice
    
    DisplayCurrentStatistics
    
    CalculateScore
    
End Sub

'===============================================================
' Procedure:    CreateWorkSheetStockAnalysis
'
' Description:  Create worksheet named Analysis - strTickerSym
'               If duplicate worksheet exists, ask user to replace
'               worksheet or cancel creation of new worksheet.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   16Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CreateWorkSheetStockAnalysis()

    Dim oSheet As Worksheet, vRet As Variant

    On Error GoTo ErrorHandler
    
    Set oSheet = Worksheets.Add
    With oSheet
        .Name = "Analysis - " & strTickerSym
        .Cells(1.1).Select
        .Activate
    End With
    
    Exit Sub
    
ErrorHandler:

    'if error due to duplicate worksheet detected
    If Err.Number = 1004 Then
        'display an options to user
        vRet = MsgBox("This file already exists.  " & _
            "Do you want to replace it?", _
            vbQuestion + vbYesNo, "Duplicate Worksheet")

        If vRet = vbYes Then
            'delete the old worksheet
            Application.DisplayAlerts = False
            Worksheets("Analysis - " & strTickerSym).Delete
            Application.DisplayAlerts = True

            'rename and activate the new worksheet
            With oSheet
                .Name = "Analysis - " & strTickerSym
                .Cells(1.1).Select
                .Activate
            End With
        Else
            'cancel the operation, delete the new worksheet
            Application.DisplayAlerts = False
            oSheet.Delete
            Application.DisplayAlerts = True
            'activate the old worksheet
            Worksheets("Analysis - " & strTickerSym).Activate
        End If

    End If

End Sub

'===============================================================
' Procedure:    AssignCellItemsStockAnalysis
'
' Description:  Assign names to cells used by each item on checklist
'               Facilitates checklist changes such as adding items
'               or changing the order.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   25Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub AssignCellItemsStockAnalysis()

    'date header
    Range("C1").Name = "DateHeader"
    Range("1:1").Name = "DateRow"
    
    'revenue checklist item
    Range("A2").Name = "ListItemRevenue"
    Range("A2:B2").Name = "LineListItemRevenue"
    Range("A2:G2").Name = "LineListItemRevenueRow"
    Range("B3").Name = "Revenue"
    Range("B4").Name = "RevenueYOYGrowth"
    Range("4:4").Name = "RevenueYOYRow"
    Range("G3:G4").Name = "RevenueCheck"
    Range("H3:H4").Name = "RevenueScore"
    
    'earnings checklist item
    Range("A5").Name = "ListItemEarnings"
    Range("A5:B5").Name = "LineListItemEarnings"
    Range("A5:G5").Name = "LineListItemEarningsRow"
    Range("B6").Name = "Earnings"
    Range("B7").Name = "EarningsYOYGrowth"
    Range("7:7").Name = "EarningsYOYRow"
    Range("G6:G7").Name = "EarningsCheck"
    
    'profit checklist item
    Range("A8").Name = "ListItemNetMargin"
    Range("A8:B8").Name = "LineListItemNetMargin"
    Range("A8:G8").Name = "LineListItemNetMarginRow"
    Range("B9").Name = "NetMargin"
    Range("9:9").Name = "NetMarginRow"
    Range("B10").Name = "NetMarginYOYGrowth"
    Range("10:10").Name = "NetMarginYOYRow"
    Range("G9:G10").Name = "ProfitsCheck"
    
    'cash flow checklist item
    Range("A11").Name = "ListItemFreeCashFlow"
    Range("A11:B11").Name = "LineListItemFreeCashFlow"
    Range("A11:G11").Name = "LineListItemFreeCashFlowRow"
    Range("B12").Name = "FreeCashFlow"
    Range("12:12").Name = "FreeCashFlowRow"
    Range("B13").Name = "FreeCashFlowYOYGrowth"
    Range("13:13").Name = "FreeCashFlowYOYRow"
    Range("G12:G13").Name = "FreeCashFlowCheck"
    
    'growth potential checklist item
    Range("A14").Name = "ListItemROE"
    Range("A14:B14").Name = "LineListItemROE"
    Range("A14:G14").Name = "LineListItemROERow"
    Range("B15").Name = "ROE"
    Range("15:15").Name = "ROERow"
    Range("B16").Name = "ROEYOYGrowth"
    Range("16:16").Name = "ROEYOYRow"
    Range("G15:G16").Name = "ROECheck"
    
    'leverage checklist item
    Range("A17").Name = "ListItemFinancialLeverage"
    Range("A17:B17").Name = "LineListItemFinancialLeverage"
    Range("A17:G17").Name = "LineListItemFinancialLeverageRow"
    
    Range("B18").Name = "LeverageRatio"
    Range("18:18").Name = "LeverageRatioRow"
    Range("B19").Name = "LeverageRatioYOYGrowth"
    Range("19:19").Name = "LeverageRatioYOYRow"
    
    Range("B20").Name = "DebtToEquity"
    Range("20:20").Name = "DebtToEquityRow"
    Range("B21").Name = "DebtToEquityYOYGrowth"
    Range("21:21").Name = "DebtToEquityYOYRow"
    
    Range("G18:G21").Name = "LeverageCheck"
    
    'liquidity checklist item
    Range("A22").Name = "ListItemQuickRatio"
    Range("A22:B22").Name = "LineListItemQuickRatio"
    Range("A22:G22").Name = "LineListItemQuickRatioRow"
    Range("B23").Name = "QuickRatio"
    Range("23:23").Name = "QuickRatioRow"
    Range("B24").Name = "QuickRatioYOYGrowth"
    Range("24:24").Name = "QuickRatioYOYRow"
    Range("G23:G24").Name = "LiquidityCheck"
    
    'red flags checklist item
    Range("A25").Name = "ListItemRedFlags"
    Range("A25:B25").Name = "LineListItemRedFlags"
    Range("A25:G25").Name = "LineListItemRedFlagsRow"
    
    Range("B26").Name = "Receivables"
    Range("26:26").Name = "ReceivablesRow"
    Range("B27").Name = "ReceivablesYOYGrowth"
    Range("27:27").Name = "ReceivablesYOYRow"
    
    Range("B28").Name = "Inventory"
    Range("28:28").Name = "InventoryRow"
    Range("B29").Name = "InventoryYOYGrowth"
    Range("29:29").Name = "InventoryYOYRow"
    
    Range("B30").Name = "SGA"
    Range("30:30").Name = "SGARow"
    Range("B31").Name = "SGAYOYGrowth"
    Range("31:31").Name = "SGAYOYRow"
    
    Range("B32").Name = "Dividend"
    Range("32:32").Name = "DividendRow"
    Range("B33").Name = "DividendYOYGrowth"
    Range("33:33").Name = "DividendYOYRow"
    
    Range("G26:G33").Name = "RedFlagsCheck"
    
    'price checklist item
    Range("A34").Name = "ListItemPrice"
    Range("A34:B34").Name = "LineListItemPrice"
    Range("A34:G34").Name = "LineListItemPriceRow"
    
    Range("B35:B37").Name = "PriceCol1"
    Range("C35:C37").Name = "PriceCol2"
    Range("E35:E37").Name = "PriceCol3"
    Range("F35:F37").Name = "PriceCol4"
    
    Range("B35").Name = "Price"
    Range("B36").Name = "TargetPrice"
    Range("B37").Name = "PriceGrowth"
    
    Range("E35").Name = "HighTarget"
    Range("E36").Name = "LowTarget"
    Range("E37").Name = "Brokers"
    
    Range("G35:G37").Name = "PriceCheck"
    
    'current statistics
    Range("A38:G38").Name = "CurrentStatsRow"
    Range("B39:B43").Name = "CurrentStatsCol1"
    Range("C39:C43").Name = "CurrentStatsCol2"
    Range("E39:E43").Name = "CurrentStatsCol3"
    Range("F39:F43").Name = "CurrentStatsCol4"
    
    Range("B39").Name = "MarketCap"
    Range("B40").Name = "PETTM"
    Range("B41").Name = "EPSTTM"
    Range("B42").Name = "DivYield"
    Range("B43").Name = "RevenueTTM"
    
    Range("E39").Name = "ProfitMarginTTM"
    Range("E40").Name = "ROETTM"
    Range("E41").Name = "DebtToEquityMRQ"
    Range("E42").Name = "CurrentRatioMRQ"
    Range("E43").Name = "FreeCashFlowTTM"
    
    Range("H39:H40").Name = "TotalScore"
End Sub

'===============================================================
' Procedure:    FormatCheckListStockAnalysis
'
' Description:  Format worksheet in order to populate with
'               checklist information. Add date headers and set
'               column widths.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   16Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListStockAnalysis()

    Dim i As Integer
    
    Worksheets("Analysis - " & strTickerSym).Activate
    
    Range("A:A").ColumnWidth = 5
    Range("B:B").ColumnWidth = COLUMN_WIDTH_CHECKLIST_ITEM
    Range("C:C").ColumnWidth = COLUMN_WIDTH_ANNUAL_DATA
    Range("D:D").ColumnWidth = COLUMN_WIDTH_ANNUAL_DATA
    Range("E:E").ColumnWidth = COLUMN_WIDTH_ANNUAL_DATA
    Range("F:F").ColumnWidth = COLUMN_WIDTH_ANNUAL_DATA
    
    ActiveWindow.DisplayGridlines = False
    
    'format year line
    Range("DateRow").Font.Bold = True
    Range("DateRow").HorizontalAlignment = xlCenter
    
    For i = 0 To (iYearsAvailableIncome - 1)
        Range("DateHeader").Offset(0, i) = strYear(i)
    Next i

    FormatCheckListRevenue
    FormatCheckListEarnings
    FormatCheckListNetMargin
    FormatCheckListCashFlow
    FormatCheckListROE
    FormatCheckListFinancialLeverage
    FormatCheckListLiquidity
    FormatCheckListRedFlags
    FormatCheckListPrice
    FormatCurrentStats
    FormatTotalScore
    
End Sub

'===============================================================
' Procedure:    FormatCheckListRevenue
'
' Description:  Format cells for revenue checklist item
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   27Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListRevenue()

    Dim i As Integer
    
    Range("ListItemRevenue").Font.Bold = True
    Range("LineListItemRevenue").Merge
    
    With Range("LineListItemRevenueRow")
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
       
    Range("Revenue").HorizontalAlignment = xlLeft
        
    With Range("RevenueYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("RevenueYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    With Range("RevenueCheck")
        .Merge
        .Font.Name = "Wingdings 2"
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Range("RevenueScore")
        .Merge
        .Font.Size = 20
        .Font.ColorIndex = FONT_COLOR_BLUE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub

'===============================================================
' Procedure:    FormatCheckListEarnings
'
' Description:  Format cells for earnings checklist item
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   27Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListEarnings()

    Range("ListItemEarnings").Font.Bold = True
    Range("LineListItemEarnings").Merge
    
    With Range("LineListItemEarningsRow")
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    Range("Earnings").HorizontalAlignment = xlLeft
    
    With Range("EarningsYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("EarningsYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    With Range("EarningsCheck")
        .Merge
        .Font.Name = "Wingdings 2"
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub

'===============================================================
' Procedure:    FormatCheckListNetMargin
'
' Description:  Format cells for profits checklist item
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   27Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListNetMargin()

    Range("ListItemNetMargin").Font.Bold = True
    Range("LineListItemNetMargin").Merge
    
    With Range("LineListItemNetMarginRow")
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    Range("NetMargin").HorizontalAlignment = xlLeft
    Range("NetMarginRow").NumberFormat = "0.0%"
    
    With Range("NetMarginYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("NetMarginYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    With Range("ProfitsCheck")
        .Merge
        .Font.Name = "Wingdings 2"
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub

'===============================================================
' Procedure:    FormatCheckListCashFlow
'
' Description:  Format cells for cash flow checklist item
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   10Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListCashFlow()

    Range("ListItemFreeCashFlow").Font.Bold = True
    Range("LineListItemFreeCashFlow").Merge
    
    With Range("LineListItemFreeCashFlowRow")
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    Range("FreeCashFlow").HorizontalAlignment = xlLeft
    
    With Range("FreeCashFlowYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("FreeCashFlowYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    With Range("FreeCashFlowCheck")
        .Merge
        .Font.Name = "Wingdings 2"
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub

'===============================================================
' Procedure:    FormatCheckListROE
'
' Description:  Format cells for ROE checklist item
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   10Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListROE()

    Range("ListItemROE").Font.Bold = True
    Range("LineListItemROE").Merge
    
    With Range("LineListItemROERow")
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    Range("ROE").HorizontalAlignment = xlLeft
    Range("ROERow").NumberFormat = "0.0%"
    
    With Range("ROEYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("ROEYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    With Range("ROECheck")
        .Merge
        .Font.Name = "Wingdings 2"
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub

'===============================================================
' Procedure:    FormatCheckListFinancialLeverage
'
' Description:  Format cells for liabilities checklist item
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   12Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListFinancialLeverage()

    Range("ListItemFinancialLeverage").Font.Bold = True
    Range("LineListItemFinancialLeverage").Merge
    
    With Range("LineListItemFinancialLeverageRow")
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    'leverage ratio
    Range("LeverageRatio").HorizontalAlignment = xlLeft
    Range("LeverageRatioRow").NumberFormat = "0.00"
    
    With Range("LeverageRatioYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("LeverageRatioYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    'debt to equity
    Range("DebtToEquity").HorizontalAlignment = xlLeft
    Range("DebtToEquityRow").NumberFormat = "0.0%"
    
    With Range("DebtToEquityYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("DebtToEquityYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    'item check
    With Range("LeverageCheck")
        .Merge
        .Font.Name = "Wingdings 2"
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub

'===============================================================
' Procedure:    FormatCheckListLiquidity
'
' Description:  Format cells for ROE checklist item
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   10Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListLiquidity()

    Range("ListItemQuickRatio").Font.Bold = True
    Range("LineListItemQuickRatio").Merge
    
    With Range("LineListItemQuickRatioRow")
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    Range("QuickRatio").HorizontalAlignment = xlLeft
    Range("QuickRatioRow").NumberFormat = "0.0"
    
    With Range("QuickRatioYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("QuickRatioYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    With Range("LiquidityCheck")
        .Merge
        .Font.Name = "Wingdings 2"
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub

'===============================================================
' Procedure:    FormatCheckListRedFlags
'
' Description:  Format cells for liabilities checklist item
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   12Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListRedFlags()

    Range("ListItemRedFlags").Font.Bold = True
    Range("LineListItemRedFlags").Merge
    
    With Range("LineListItemRedFlagsRow")
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    'receivables
    Range("Receivables").HorizontalAlignment = xlLeft
    Range("ReceivablesRow").NumberFormat = "0.0%"
    
    With Range("ReceivablesYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("ReceivablesYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    'inventory
    Range("Inventory").HorizontalAlignment = xlLeft
    Range("InventoryRow").NumberFormat = "0.0%"
    
    With Range("InventoryYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("InventoryYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    'SGA
    Range("SGA").HorizontalAlignment = xlLeft
    Range("SGARow").NumberFormat = "0.0%"
    
    With Range("SGAYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("SGAYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    'dividend
    Range("dividend").HorizontalAlignment = xlLeft
    Range("dividendRow").NumberFormat = "0.00"
    
    With Range("dividendYOYGrowth")
        .HorizontalAlignment = xlRight
        .Offset(0, iYearsAvailableIncome).HorizontalAlignment = xlCenter
        .Offset(0, iYearsAvailableIncome) = STR_NO_DATA
    End With
    
    With Range("dividendYOYRow")
        .Font.Italic = True
        .NumberFormat = "0.0%"
    End With
    
    'item check
    With Range("RedFlagsCheck")
        .Merge
        .Font.Name = "Wingdings 2"
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub

'===============================================================
' Procedure:    FormatCheckListPrice
'
' Description:  Format cells for price checklist item
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   15Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCheckListPrice()

    Range("ListItemPrice").Font.Bold = True
    Range("LineListItemPrice").Merge
    
    With Range("LineListItemPriceRow")
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    Range("PriceCol1").HorizontalAlignment = xlRight
    Range("PriceCol2").HorizontalAlignment = xlCenter
    Range("PriceCol3").HorizontalAlignment = xlRight
    Range("PriceCol4").HorizontalAlignment = xlCenter
    
    Range("Price").Offset(0, 1).NumberFormat = "0.00"
    Range("TargetPrice").Offset(0, 1).NumberFormat = "0.00"
    Range("HighTarget").Offset(0, 1).NumberFormat = "0.00"
    Range("LowTarget").Offset(0, 1).NumberFormat = "0.00"
    Range("PriceGrowth").Offset(0, 1).NumberFormat = "0.0%"
    
    With Range("PriceCheck")
        .Merge
        .Font.Name = "Wingdings 2"
        .Font.Size = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub

'===============================================================
' Procedure:    FormatCurrentStats
'
' Description:  Format cells for current statistics items
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   23Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatCurrentStats()

    With Range("CurrentStatsRow")
        .Font.Bold = True
        .Merge
        .Interior.ColorIndex = COLOR_LIGHT_BLUE
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    Range("CurrentStatsCol1").HorizontalAlignment = xlRight
    Range("CurrentStatsCol2").HorizontalAlignment = xlCenter
    Range("CurrentStatsCol3").HorizontalAlignment = xlRight
    Range("CurrentStatsCol4").HorizontalAlignment = xlCenter
    
    Range("ProfitMarginTTM").Offset(0, 1).NumberFormat = "0.00%"
    Range("ROETTM").Offset(0, 1).NumberFormat = "0.00%"

End Sub

'===============================================================
' Procedure:    FormatTotalScore
'
' Description:  Format cells for total score
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   10Dec15 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatTotalScore()

    With Range("TotalScore")
        .Merge
        .Font.Size = 20
        .Font.ColorIndex = FONT_COLOR_BLUE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

End Sub

'===============================================================
' Procedure:    CalculateScore
'
' Description:  Calculate total score from each category
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   10Dec15 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateScore()

    Range("TotalScore").FormulaR1C1 = "=SUM(R[-38]C:R[-2]C)"
End Sub
