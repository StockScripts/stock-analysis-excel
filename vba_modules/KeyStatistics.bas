Attribute VB_Name = "KeyStatistics"
Option Explicit

Global dblTargetPrice As Double
Global dblCurrentPrice As Double
Global vHighTarget As Variant
Global vLowTarget As Variant
Global vBrokers As Variant
Global strSummaryMarketCap As String
Global strSummaryPE As String
Global strSummaryEPS As String
Global strSummaryDivYield As String
Global strSummaryRevenue As String
Global strSummaryProfitMargin As String
Global strSummaryROE As String
Global strSummaryDebtToEquity As String
Global strSummaryCurrentRatio As String
Global strSummaryFreeCashFlow As String

Sub CreateKeyStatistics()

    CreateKeyStatisticsSheet
    GetKeyStatistics
    FormatKeyStatisticsSheet

End Sub

Sub CreateKeyStatisticsSheet()

    Dim objSheet As Worksheet
    Dim vRet As Variant

    On Error GoTo ErrorHandler
    
    Set objSheet = Worksheets.Add
    With objSheet
        .Name = "Summary - " & strTickerSym
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
            Worksheets("Balance Sheet - " & strTickerSym).Delete
            Application.DisplayAlerts = True

            'rename and activate the new worksheet
            With objSheet
                .Name = "Balance Sheet - " & strTickerSym
                .Cells(1.1).Select
                .Activate
            End With
        Else
            'cancel the operation, delete the new worksheet
            Application.DisplayAlerts = False
            objSheet.Delete
            Application.DisplayAlerts = True
            
            'activate the old worksheet
            Worksheets("Balance Sheet - " & strTickerSym).Activate
        End If

    End If
    
End Sub

Sub GetKeyStatistics()

    Dim strTicker As String
    On Error GoTo ErrorHandler
    
    strTicker = LCase(strTickerSym)
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://finance.yahoo.com/q?s=" & strTicker & "&ql=1", Destination:=Range("$A$1"))
        .Name = "q?s=" & strTicker & "&ql&ql=1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = """table1"",""table2"""
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With

    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://finance.yahoo.com/q/ks?s=" & strTicker & "+Key+Statistics", Destination:=Range( _
        "$A$18"))
        .Name = "ks?s=" & strTicker & "+Key+Statistics"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "9,12,14,16,18,20,22"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://finance.yahoo.com/q/ao?s=" & strTicker & "+Analyst+Opinion", Destination:= _
        Range("$A$65"))
        .Name = "ao?s=" & strTicker & "+Analyst+Opinion"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "10"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
    Exit Sub
    
ErrorHandler:
    'if unable to open web page
    If Err.Number = 1004 Then
        MsgBox "Unable to obtain stock information." & _
        vbNewLine & "  - Please verify ticker symbol is correct." & _
        vbNewLine & "  - Please check internet connection.", vbExclamation
        
        End
    End If
    
End Sub

Sub FormatKeyStatisticsSheet()

    GetMarketCap
    GetPETTM
    GetEPSTTM
    GetDivYield
    GetRevenueTTM
    GetProfitMarginTTM
    GetROETTM
    GetDebtToEquityMRQ
    GetCurrentRatioMRQ
    GetFreeCashFlowTTM
    
    GetTargetPrice
    GetCurrentPrice
    GetHighTarget
    GetLowTarget
    GetBrokers

End Sub

Sub GetTargetPrice()

    Dim strTargetEst As String
    
    On Error GoTo ErrorHandler
    
    Sheets("Summary - " & strTickerSym).Activate
    
    'account item term to search for in balance sheet
    strTargetEst = "1y Target Est:"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strTargetEst, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblTargetPrice = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    dblTargetPrice = 0
    
End Sub

Sub GetCurrentPrice()

    Dim strCurrentPrice As String
    
    'account item term to search for in balance sheet
    strCurrentPrice = "Prev Close:"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strCurrentPrice, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblCurrentPrice = Selection.Offset(0, 1).Value
    
End Sub

Sub GetHighTarget()

    Dim strHighTarget As String
    
    On Error Resume Next
    
    'item term to search for
    strHighTarget = "High Target:"
    
    Columns("A:A").Select
    Selection.Find(What:=strHighTarget, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    If Err Then
        vHighTarget = "N/A"
    Else
        vHighTarget = Selection.Offset(0, 1).Value
    End If
          
End Sub

Sub GetBrokers()

    Dim strBrokers As String
    
    On Error GoTo ErrorHandler
    
    'item term to search for
    strBrokers = "No. of Brokers:"
    
    Columns("A:A").Select
    Selection.Find(What:=strBrokers, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    vBrokers = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    vBrokers = "N/A"
    
End Sub

Sub GetLowTarget()

    Dim strLowTarget As String
    
    On Error GoTo ErrorHandler
    
    'item term to search for
    strLowTarget = "Low Target:"
    
    Columns("A:A").Select
    Selection.Find(What:=strLowTarget, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    vLowTarget = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    vLowTarget = "N/A"
    
End Sub

Sub GetMarketCap()

    Dim strMarketCap As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for
    strMarketCap = "Market Cap:"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strMarketCap, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryMarketCap = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryMarketCap = "N/A"
    
End Sub

Sub GetPETTM()

    Dim strPEttm As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for
    strPEttm = "P/E (ttm):"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strPEttm, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryPE = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryPE = "N/A"
End Sub

Sub GetEPSTTM()

    Dim strEPSttm As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    strEPSttm = "EPS (ttm):"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strEPSttm, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryEPS = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryEPS = "N/A"
    
End Sub

Sub GetDivYield()

    Dim strDivYield As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    strDivYield = "Div & Yield:"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strDivYield, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryDivYield = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryDivYield = "N/A"
    
End Sub

Sub GetRevenueTTM()

    Dim strRevenueTTM As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    strRevenueTTM = "Revenue (ttm):"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strRevenueTTM, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryRevenue = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryRevenue = "N/A"
    
End Sub

Sub GetProfitMarginTTM()

    Dim strProfitMarginTTM As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    strProfitMarginTTM = "Profit Margin (ttm):"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strProfitMarginTTM, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryProfitMargin = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryProfitMargin = "N/A"
    
End Sub

Sub GetROETTM()

    Dim strROETTM As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    strROETTM = "Return on Equity (ttm):"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strROETTM, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryROE = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryROE = "N/A"
    
End Sub

Sub GetDebtToEquityMRQ()

    Dim strDebtToEquityMRQ As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    strDebtToEquityMRQ = "Total Debt/Equity (mrq):"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strDebtToEquityMRQ, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryDebtToEquity = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryDebtToEquity = "N/A"
    
End Sub

Sub GetCurrentRatioMRQ()

    Dim strCurrentRatioMRQ As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    strCurrentRatioMRQ = "Current Ratio (mrq):"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strCurrentRatioMRQ, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryCurrentRatio = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryCurrentRatio = "N/A"
    
End Sub

Sub GetFreeCashFlowTTM()

    Dim strFreeCashFlowTTM As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    strFreeCashFlowTTM = "Levered Free Cash Flow (ttm):"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=strFreeCashFlowTTM, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    strSummaryFreeCashFlow = Selection.Offset(0, 1).Value
    
    Exit Sub
    
ErrorHandler:

    strSummaryFreeCashFlow = "N/A"
    
End Sub


