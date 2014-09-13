Attribute VB_Name = "StatementBalanceSheet"
Option Explicit

Global dblReceivables(0 To 4) As Double
Global dblInventory(0 To 4) As Double
Global dblCurrentAssets(0 To 4) As Double
Global dblCurrentLiabilities(0 To 4) As Double
Global dblLongTermDebt(0 To 4) As Double
Global dblEquity(0 To 4) As Double
Global iYear(0 To 4) As Integer

'===============================================================
' Procedure:    CreateStatementBalanceSheet
'
' Description:  Call procedures to create Balance Sheet worksheet,
'               acquire data from msnmoney.com, and format
'               worksheet.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   09Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CreateStatementBalanceSheet()

    CreateWorkSheetBalanceSheet
    GetAnnualDataBalanceSheet
    FormatStatementBalanceSheet

End Sub

'===============================================================
' Procedure:    CreateWorkSheetBalanceSheet
'
' Description:  Create worksheet named Balance Sheet - strTickerSym
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
'Rev History:   09Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CreateWorkSheetBalanceSheet()

    Dim objSheet As Worksheet
    Dim vRet As Variant

    On Error GoTo ErrorHandler
    
    Set objSheet = Worksheets.Add
    With objSheet
        .Name = "Balance Sheet - " & strTickerSym
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

'===============================================================
' Procedure:    GetAnnualDataBalanceSheet
'
' Description:  Get annual Balance Sheet statement from msnmoney.com
'
' Author:       Janice Laset Parkerson
'
' Notes:        Code generated using recorded macro
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   09Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetAnnualDataBalanceSheet()

    On Error GoTo ErrorHandler
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://investing.money.msn.com/investments/stock-balance-sheet/?symbol=us%3A" & strTickerSym & "&stmtView=Ann" _
        , Destination:=Range("$A$1"))
        .Name = "?symbol=us%3ASLP&stmtView=Ann"
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
        .WebTables = "2"
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

'===============================================================
' Procedure:    FormatStatementBalanceSheet
'
' Description:  Get info required from balance sheet and highlight
'               items
'               - receivables
'               - inventory
'               - current assets
'               - current liabilities
'               - long term debt
'               - equity
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatStatementBalanceSheet()

    Sheets("Balance Sheet - " & strTickerSym).Activate
    
    GetYears
    GetReceivables
    GetInventory
    GetCurrentAssets
    GetCurrentLiabilities
    GetLongTermDebt
    GetEquity
    
End Sub

'===============================================================
' Procedure:    GetYears
'
' Description:  Get year values for financial report
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetYears()

    iYear(0) = ActiveSheet.Range("B1").value
    iYear(1) = ActiveSheet.Range("C1").value
    iYear(2) = ActiveSheet.Range("D1").value
    iYear(3) = ActiveSheet.Range("E1").value
    iYear(4) = ActiveSheet.Range("F1").value

End Sub

'===============================================================
' Procedure:    GetReceivables
'
' Description:  Find Receivables information in balance sheet
'               and get annual data
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetReceivables()

    Dim Receivables As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    Receivables = "Receivables"
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=Receivables, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblReceivables(0) = Selection.Offset(0, 1).value
    dblReceivables(1) = Selection.Offset(0, 2).value
    dblReceivables(2) = Selection.Offset(0, 3).value
    dblReceivables(3) = Selection.Offset(0, 4).value
    dblReceivables(4) = Selection.Offset(0, 5).value
        
    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Receivables information."
   
    dblReceivables(0) = 0
    dblReceivables(1) = 0
    dblReceivables(2) = 0
    dblReceivables(3) = 0
    dblReceivables(4) = 0

End Sub

'===============================================================
' Procedure:    GetInventory
'
' Description:  Find Inventory information in balance sheet
'               and get annual data
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetInventory()

    Dim Inventory As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    Inventory = "Inventories"
    
    'find inventory account item
    Columns("A:A").Select
    Selection.Find(What:=Inventory, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblInventory(0) = Selection.Offset(0, 1).value
    dblInventory(1) = Selection.Offset(0, 2).value
    dblInventory(2) = Selection.Offset(0, 3).value
    dblInventory(3) = Selection.Offset(0, 4).value
    dblInventory(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
        
    Exit Sub
        
ErrorHandler:
   MsgBox "No Inventory information."
   
   dblInventory(0) = 0
   dblInventory(1) = 0
   dblInventory(2) = 0
   dblInventory(3) = 0
   dblInventory(4) = 0
   
End Sub

'===============================================================
' Procedure:    GetCurrentAssets
'
' Description:  Find current assets information in balance sheet
'               and get annual data
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetCurrentAssets()

    Dim CurrentAssets As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    CurrentAssets = "Total Current Assets"
        
    'find current assets account item
    Columns("A:A").Select
    Selection.Find(What:=CurrentAssets, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblCurrentAssets(0) = Selection.Offset(0, 1).value
    dblCurrentAssets(1) = Selection.Offset(0, 2).value
    dblCurrentAssets(2) = Selection.Offset(0, 3).value
    dblCurrentAssets(3) = Selection.Offset(0, 4).value
    dblCurrentAssets(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Current Assets information."
    
    dblCurrentAssets(0) = 0
    dblCurrentAssets(0) = 0
    dblCurrentAssets(0) = 0
    dblCurrentAssets(0) = 0
    dblCurrentAssets(0) = 0

End Sub

'===============================================================
' Procedure:    GetCurrentLiabilities
'
' Description:  Find current liabilities information in balance sheet
'               and get annual data
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetCurrentLiabilities()

    Dim CurrentLiabilities As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    CurrentLiabilities = "Total Current Liabilities"
        
    'find current liabilities account item
    Columns("A:A").Select
    Selection.Find(What:=CurrentLiabilities, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblCurrentLiabilities(0) = Selection.Offset(0, 1).value
    dblCurrentLiabilities(1) = Selection.Offset(0, 2).value
    dblCurrentLiabilities(2) = Selection.Offset(0, 3).value
    dblCurrentLiabilities(3) = Selection.Offset(0, 4).value
    dblCurrentLiabilities(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font

    Exit Sub
    
ErrorHandler:
    MsgBox "No Current Liabilities information."
    
    dblCurrentLiabilities(0) = 0
    dblCurrentLiabilities(1) = 0
    dblCurrentLiabilities(2) = 0
    dblCurrentLiabilities(3) = 0
    dblCurrentLiabilities(4) = 0
    
End Sub

'===============================================================
' Procedure:    GetLongTermDebt
'
' Description:  Find long term debt information in balance sheet
'               and get annual data
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetLongTermDebt()

    Dim LongTermDebt As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    LongTermDebt = "Lt Debt and Capital Lease Obligation"
    
    'find long term debt account item
    Columns("A:A").Select
    Selection.Find(What:=LongTermDebt, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblLongTermDebt(0) = Selection.Offset(0, 1).value
    dblLongTermDebt(1) = Selection.Offset(0, 2).value
    dblLongTermDebt(2) = Selection.Offset(0, 3).value
    dblLongTermDebt(3) = Selection.Offset(0, 4).value
    dblLongTermDebt(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Long Term Debt information."
    
    dblLongTermDebt(0) = 0
    dblLongTermDebt(1) = 0
    dblLongTermDebt(2) = 0
    dblLongTermDebt(3) = 0
    dblLongTermDebt(4) = 0
    
End Sub

'===============================================================
' Procedure:    GetEquity
'
' Description:  Find equity information in balance sheet
'               and get annual data
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetEquity()

    Dim Equity As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    Equity = "Total Equity"
       
    'find equity account item
    Columns("A:A").Select
    Selection.Find(What:=Equity, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblEquity(0) = Selection.Offset(0, 1).value
    dblEquity(1) = Selection.Offset(0, 2).value
    dblEquity(2) = Selection.Offset(0, 3).value
    dblEquity(3) = Selection.Offset(0, 4).value
    dblEquity(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Equity information."
    
    dblEquity(0) = 0
    dblEquity(1) = 0
    dblEquity(2) = 0
    dblEquity(3) = 0
    dblEquity(4) = 0
    
End Sub
