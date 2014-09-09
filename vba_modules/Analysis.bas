Attribute VB_Name = "Analysis"
Option Explicit

'Public TickerSym As String
'Public Const RedFont = 3
'Public Const GreenFont = 10
'Public Const OrangeFont = 46

'Sub FindStocks()
'    FindStocksForm.Show
'End Sub


'Sub StockAnalysis()
'    TickerSymForm.Show
'End Sub

'Sub AnalyzeStock()
 '   BalanceSheetStatement
  '  CashFlowStatement
   ' IncomeStatement
    'StocksChecklist
'End Sub

'Callback for Button1 onAction
Sub FindStocks(control As IRibbonControl)
    FindStocksForm.Show
End Sub

'Callback for Button2 onAction
Sub StockAnalysis(control As IRibbonControl)
    TickerSymForm.Show
End Sub
