Attribute VB_Name = "Analysis"
Option Explicit

'Globals used throughout project
Global strTickerSym As String
Global Const FONT_COLOR_RED = 3
Global Const FONT_COLOR_GREEN = 10
Global Const FONT_COLOR_ORANGE = 46
Global Const FONT_COLOR_BLUE = 5

Global Const STR_NO_DATA = "---"    'indicates no data obtained from statement

Global Const CHECK_MARK = "P"
Global Const X_MARK = "O"

Global Const YEARS_MAX = 4  '4 years - used in 0 based for loops

Public Enum Result
    PASS
    FAIL
End Enum

'===============================================================
' Procedure:    AnalyzeStock
'
' Description:  Call procedures to get data from MSN Money site.
'               Create worksheet for Balance Sheet, Cash Flow
'               Statement, and Income Statement.
'               Extract required information and generate stocks
'               checklist to be used to analyze company.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   09Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub AnalyzeStock()

    ImportFinancialData
    CreateStatementIncome
    CreateStatementBalanceSheet
    CreateStatementCashFlow
    CreateKeyStatistics
    
    ie.Quit
    
    CreateChecklistStockAnalysis
        
End Sub

'===============================================================
' Procedure:    FindStocks(control As IRibbonControl)
'
' Description:  Callback for Find Stocks button
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   09Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FindStocks(control As IRibbonControl)

    FormFindStocks.Show
    
End Sub

'===============================================================
' Procedure:    FindStocks(control As IRibbonControl)
'
' Description:  Callback for Stock Analysis button
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   09Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub StockAnalysis(control As IRibbonControl)

    FormTickerSym.Show
    
End Sub

'===============================================================
' Procedure:    CalculateYOYGrowth
'
' Description:  Calculate year over year growth between recent
'               and past year
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   RecentYear, PastYear
'
' Returns:      Year over year growth between past and recent year
'
'Rev History:   09Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function CalculateYOYGrowth(RecentYear, PastYear)
    On Error Resume Next
    If PastYear = 0 Then
        CalculateYOYGrowth = 0
    Else
        CalculateYOYGrowth = (RecentYear - PastYear) / Abs(PastYear)
    End If
    
End Function
