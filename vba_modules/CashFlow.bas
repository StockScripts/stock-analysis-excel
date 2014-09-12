Attribute VB_Name = "CashFlow"
Option Explicit

Global dblOpCashFlow(0 To 4) As Double
Global dblFreeCashFlow(0 To 4) As Double

'===============================================================
' Procedure:    CreateStatementCashFlow
'
' Description:  Call procedures to create Cash Flow statement
'               worksheet, acquire data from msnmoney.com, and format
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
'Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CreateStatementCashFlow()

    CreateWorksheetCashFlow
    GetAnnualDataCashFlow
    FormatStatementCashFlow

End Sub

'===============================================================
' Procedure:    CreateWorksheetCashFlow
'
' Description:  Create worksheet named Cash Flow - strTickerSym
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
'Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CreateWorksheetCashFlow()

    Dim oSheet As Worksheet, vRet As Variant

    On Error GoTo ErrorHandler
    
    Set oSheet = Worksheets.Add
    With oSheet
        .Name = "Cash Flow - " & strTickerSym
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
            Worksheets("Cash Flow - " & strTickerSym).Delete
            Application.DisplayAlerts = True

            'rename and activate the new worksheet
            With oSheet
                .Name = "Cash Flow - " & strTickerSym
                .Cells(1.1).Select
                .Activate
            End With
        Else
            'cancel the operation, delete the new worksheet
            Application.DisplayAlerts = False
            oSheet.Delete
            Application.DisplayAlerts = True
            'activate the old worksheet
            Worksheets("Cash Flow - " & strTickerSym).Activate
        End If

    End If
    
End Sub

'===============================================================
' Procedure:    GetAnnualDataBalanceSheet
'
' Description:  Get annual Cash Flow statement from msnmoney.com
'
' Author:       Janice Laset Parkerson
'
' Notes:        Code generated using recorded macro
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetAnnualDataCashFlow()

    On Error GoTo ErrorHandler
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://investing.money.msn.com/investments/stock-cash-flow/?symbol=" & strTickerSym & "", _
        Destination:=Range("$A$1"))
        .Name = "?symbol=slp"
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
' Procedure:    FormatStatementCashFlow
'
' Description:  Get info required from balance sheet and highlight
'               items
'               - operating cash flow
'               - free cash flow
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
Sub FormatStatementCashFlow()
    
    Sheets("Cash Flow - " & strTickerSym).Activate
    
    GetOpCashFlow
    GetFreeCashFlow
        
End Sub

'===============================================================
' Procedure:    GetOpCashFlow
'
' Description:  Find operating cash flow information in cash flow
'               statement and get annual data
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
Sub GetOpCashFlow()

    Dim OpCashFlow As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    OpCashFlow = "Cash Flow from Operating Activities"
    
    'find operating cash flow account item
    Columns("A:A").Select
    Selection.Find(What:=OpCashFlow, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblOpCashFlow(0) = Selection.Offset(0, 1).value
    dblOpCashFlow(1) = Selection.Offset(0, 2).value
    dblOpCashFlow(2) = Selection.Offset(0, 3).value
    dblOpCashFlow(3) = Selection.Offset(0, 4).value
    dblOpCashFlow(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
   MsgBox "No Operating Cash Flow information."
   
   dblOpCashFlow(0) = 0
   dblOpCashFlow(1) = 0
   dblOpCashFlow(2) = 0
   dblOpCashFlow(3) = 0
   dblOpCashFlow(4) = 0
    
End Sub

'===============================================================
' Procedure:    GetFreeCashFlow
'
' Description:  Find free cash flow information in cash flow
'               statement and get annual data
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
Sub GetFreeCashFlow()

    Dim FreeCashFlow As String

    On Error GoTo ErrorHandler

    'account item term to search for in balance sheet
    FreeCashFlow = "Free Cash Flow"
    
    'find free cash flow account item
    Columns("A:A").Select
    Selection.Find(What:=FreeCashFlow, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblFreeCashFlow(0) = Selection.Offset(0, 1).value
    dblFreeCashFlow(1) = Selection.Offset(0, 2).value
    dblFreeCashFlow(2) = Selection.Offset(0, 3).value
    dblFreeCashFlow(3) = Selection.Offset(0, 4).value
    dblFreeCashFlow(4) = Selection.Offset(0, 5).value
    
    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
   MsgBox "No Free Cash Flow information."
    
   dblFreeCashFlow(0) = 0
   dblFreeCashFlow(1) = 0
   dblFreeCashFlow(2) = 0
   dblFreeCashFlow(3) = 0
   dblFreeCashFlow(4) = 0
End Sub
