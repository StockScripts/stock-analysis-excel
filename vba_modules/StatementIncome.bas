Attribute VB_Name = "StatementIncome"
Option Explicit

Global dblRevenue(0 To 4) As Double
Global dblSGA(0 To 4) As Double
Global dblNetIncome(0 To 4) As Double
Global dblEPS(0 To 4) As Double
Global dblDividendPerShare(0 To 4) As Double

'===============================================================
' Procedure:    CreateStatementIncome
'
' Description:  Call procedures to create Income statement worksheet,
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
'Rev History:   12Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CreateStatementIncome()

    CreateWorksheetIncome
    GetAnnualDataIncome
    FormatStatementIncome

End Sub

'===============================================================
' Procedure:    CreateWorksheetIncome
'
' Description:  Create worksheet named Income - strTickerSym
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
'Rev History:   12Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CreateWorksheetIncome()

    Dim oSheet As Worksheet, vRet As Variant

    On Error GoTo ErrorHandler
    
    Set oSheet = Worksheets.Add
    With oSheet
        .Name = "Income - " & strTickerSym
        .Cells(1.1).Select
        .Activate
    End With
    
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
            Worksheets("Income - " & strTickerSym).Delete
            Application.DisplayAlerts = True

            'rename and activate the new worksheet
            With oSheet
                .Name = "Income - " & strTickerSym
                .Cells(1.1).Select
                .Activate
            End With
        Else
            'cancel the operation, delete the new worksheet
            Application.DisplayAlerts = False
            oSheet.Delete
            Application.DisplayAlerts = True
            'activate the old worksheet
            Worksheets("Income - " & strTickerSym).Activate
        End If

    End If
End Sub

'===============================================================
' Procedure:    GetAnnualDataIncome
'
' Description:  Get annual income statement from msnmoney.com
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
Sub GetAnnualDataIncome()

    On Error GoTo ErrorHandler

    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://investing.money.msn.com/investments/stock-income-statement/?symbol=" & strTickerSym & "" _
        , Destination:=Range("$A$1"))
        .Name = "?symbol=SLP"
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
' Procedure:    FormatStatementIncome
'
' Description:  Get info required from balance sheet and highlight
'               items
'               - revenue
'               - SGA
'               - net income
'               - EPS
'               - dividend per share
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:   12Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub FormatStatementIncome()

    Sheets("Income - " & strTickerSym).Activate
    
    GetRevenue
    GetSGA
    GetNetIncome
    GetEPS
    GetDividendPerShare
    
End Sub

'===============================================================
' Procedure:    GetRevenue
'
' Description:  Find revenue information in income statement
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
' Rev History:   12Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetRevenue()

    Dim Revenue As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in income statement
    Revenue = "Total Revenue"
    
    'find revenue account item
    Columns("A:A").Select
    Selection.Find(What:=Revenue, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select

    dblRevenue(0) = Selection.Offset(0, 1).value
    dblRevenue(1) = Selection.Offset(0, 2).value
    dblRevenue(2) = Selection.Offset(0, 3).value
    dblRevenue(3) = Selection.Offset(0, 4).value
    dblRevenue(4) = Selection.Offset(0, 5).value
    
    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Revenue information."
    
    dblRevenue(0) = 0
    dblRevenue(1) = 0
    dblRevenue(2) = 0
    dblRevenue(3) = 0
    dblRevenue(4) = 0
    
End Sub

'===============================================================
' Procedure:    GetSGA
'
' Description:  Find SGA information in income statement
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
' Rev History:   12Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetSGA()

    Dim SGA As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in income statement
    SGA = "Selling"
       
    'find SGA account item
    Columns("A:A").Select
    Selection.Find(What:=SGA, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblSGA(0) = Selection.Offset(0, 1).value
    dblSGA(1) = Selection.Offset(0, 2).value
    dblSGA(2) = Selection.Offset(0, 3).value
    dblSGA(3) = Selection.Offset(0, 4).value
    dblSGA(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No SGA information."
    
    dblSGA(0) = 0
    dblSGA(1) = 0
    dblSGA(2) = 0
    dblSGA(3) = 0
    dblSGA(4) = 0
    
End Sub

'===============================================================
' Procedure:    GetNetIncome
'
' Description:  Find net income information in income statement
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
' Rev History:   12Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetNetIncome()

    Dim NetIncome As String
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in income statement
    NetIncome = "Net Income"
         
    'find net income account item
    Range("A:A").Select
    Selection.Find(What:=NetIncome, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblNetIncome(0) = Selection.Offset(0, 1).value
    dblNetIncome(1) = Selection.Offset(0, 2).value
    dblNetIncome(2) = Selection.Offset(0, 3).value
    dblNetIncome(3) = Selection.Offset(0, 4).value
    dblNetIncome(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Net Income information."
    
    dblNetIncome(0) = 0
    dblNetIncome(1) = 0
    dblNetIncome(2) = 0
    dblNetIncome(3) = 0
    dblNetIncome(4) = 0
    
End Sub

'===============================================================
' Procedure:    GetEPS
'
' Description:  Find EPS information in income statement
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
' Rev History:   12Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetEPS()

    Dim EPS As String

    'account item term to search for in income statement
    EPS = "Diluted EPS"
    
    On Error GoTo ErrorHandler
    
    'find EPS account item
    Columns("A:A").Select
    Selection.Find(What:=EPS, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblEPS(0) = Selection.Offset(0, 1).value
    dblEPS(1) = Selection.Offset(0, 2).value
    dblEPS(2) = Selection.Offset(0, 3).value
    dblEPS(3) = Selection.Offset(0, 4).value
    dblEPS(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No EPS information."
    
    dblEPS(0) = 0
    dblEPS(1) = 0
    dblEPS(2) = 0
    dblEPS(3) = 0
    dblEPS(4) = 0

End Sub

'===============================================================
' Procedure:    GetDividendPerShare
'
' Description:  Find dividend per share information in income statement
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
' Rev History:   12Sept2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetDividendPerShare()

    Dim DividendPerShare As String

    'account item term to search for in income statement
    DividendPerShare = "Dividend Per Share"
    
    On Error GoTo ErrorHandler
    
    'find dividend per share account item
    Columns("A:A").Select
    Selection.Find(What:=DividendPerShare, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    dblDividendPerShare(0) = Selection.Offset(0, 1).value
    dblDividendPerShare(1) = Selection.Offset(0, 2).value
    dblDividendPerShare(2) = Selection.Offset(0, 3).value
    dblDividendPerShare(3) = Selection.Offset(0, 4).value
    dblDividendPerShare(4) = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Dividend Per Share information."
    
    dblDividendPerShare(0) = 0
    dblDividendPerShare(1) = 0
    dblDividendPerShare(2) = 0
    dblDividendPerShare(3) = 0
    dblDividendPerShare(4) = 0

End Sub
