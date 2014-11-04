Attribute VB_Name = "CurrentStatistics"
Option Explicit

Sub DisplayCurrentStatistics()

    Range("CurrentStatsRow") = "Current Statistics"
    
    Range("MarketCap") = "Market Cap"
    Range("MarketCap").Offset(0, 1) = strSummaryMarketCap
    
    Range("PETTM") = "P/E (ttm)"
    Range("PETTM").Offset(0, 1) = strSummaryPE
    
    Range("EPSTTM") = "EPS (ttm)"
    Range("EPSTTM").Offset(0, 1) = strSummaryEPS
    
    Range("DivYield") = "Div Yield"
    Range("DivYield").Offset(0, 1) = strSummaryDivYield
    
    Range("RevenueTTM") = "Revenue (ttm)"
    Range("RevenueTTM").Offset(0, 1) = strSummaryRevenue
    
    Range("ProfitMarginTTM") = "Profit Margin (ttm)"
    Range("ProfitMarginTTM").Offset(0, 1) = strSummaryProfitMargin
    
    Range("ROETTM") = "ROE (ttm)"
    Range("ROETTM").Offset(0, 1) = strSummaryROE
    
    Range("DebtToEquityMRQ") = "Total Debt To Equity (mrq)"
    Range("DebtToEquityMRQ").Offset(0, 1) = strSummaryDebtToEquity
    
    Range("CurrentRatioMRQ") = "Current Ratio (mrq)"
    Range("CurrentRatioMRQ").Offset(0, 1) = strSummaryCurrentRatio
    
    Range("FreeCashFlowTTM") = "Free Cash Flow (ttm)"
    Range("FreeCashFlowTTM").Offset(0, 1) = strSummaryFreeCashFlow

End Sub
