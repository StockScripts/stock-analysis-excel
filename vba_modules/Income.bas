Attribute VB_Name = "Income"
Option Explicit

Public Revenue1 As Double
Public Revenue2 As Double
Public Revenue3 As Double
Public Revenue4 As Double
Public Revenue5 As Double

Public SGA1 As Double
Public SGA2 As Double
Public SGA3 As Double
Public SGA4 As Double
Public SGA5 As Double

Public NetIncome1 As Double
Public NetIncome2 As Double
Public NetIncome3 As Double
Public NetIncome4 As Double
Public NetIncome5 As Double

Public EPS1 As Double
Public EPS2 As Double
Public EPS3 As Double
Public EPS4 As Double
Public EPS5 As Double

Public DividendPerShare1 As Double
Public DividendPerShare2 As Double
Public DividendPerShare3 As Double
Public DividendPerShare4 As Double
Public DividendPerShare5 As Double


'Create Income Statement with data from msnmoney.com
Sub IncomeStatement()

    CreateIncomeStatement
    GetIncomeStatement
    FormatIncomeStatement

End Sub

Sub CreateIncomeStatement()

    Dim oSheet As Worksheet, vRet As Variant

    On Error GoTo ErrorHandler
    
    Set oSheet = Worksheets.Add
    With oSheet
        .Name = "Income - " & TickerSym
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
            Worksheets("Income - " & TickerSym).Delete
            Application.DisplayAlerts = True

            'rename and activate the new worksheet
            With oSheet
                .Name = "Income - " & TickerSym
                .Cells(1.1).Select
                .Activate
            End With
        Else
            'cancel the operation, delete the new worksheet
            Application.DisplayAlerts = False
            oSheet.Delete
            Application.DisplayAlerts = True
            'activate the old worksheet
            Worksheets("Income - " & TickerSym).Activate
        End If

    End If
End Sub

'Gets annual Income Statement from msnmoney.com
Sub GetIncomeStatement()

    On Error GoTo ErrorHandler

    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://investing.money.msn.com/investments/stock-income-statement/?symbol=" & TickerSym & "" _
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

'Assign account names to cells in Income Statement
Sub FormatIncomeStatement()

    Sheets("Income - " & TickerSym).Activate
    
    GetRevenue
    GetSGA
    GetNetIncome
    GetEPS
    GetDividendPerShare
    
End Sub

Sub GetRevenue()

    Dim Revenue As String
    
    Revenue = "Total Revenue"
    
    On Error GoTo ErrorHandler
    
    'Revenue
    Columns("A:A").Select
    Selection.Find(What:=Revenue, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select

    Revenue1 = Selection.Offset(0, 1).value
    Revenue2 = Selection.Offset(0, 2).value
    Revenue3 = Selection.Offset(0, 3).value
    Revenue4 = Selection.Offset(0, 4).value
    Revenue5 = Selection.Offset(0, 5).value
    
    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Revenue information."
    
    Revenue1 = 0
    Revenue2 = 0
    Revenue3 = 0
    Revenue4 = 0
    Revenue5 = 0
    
End Sub

Sub GetSGA()

    Dim SGA As String
    
    SGA = "Selling"
    
    On Error GoTo ErrorHandler
    
    'SGA
    Columns("A:A").Select
    Selection.Find(What:=SGA, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    SGA1 = Selection.Offset(0, 1).value
    SGA2 = Selection.Offset(0, 2).value
    SGA3 = Selection.Offset(0, 3).value
    SGA4 = Selection.Offset(0, 4).value
    SGA5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No SGA information."
    
    SGA1 = 0
    SGA2 = 0
    SGA3 = 0
    SGA4 = 0
    SGA5 = 0
    
End Sub

Sub GetNetIncome()

    Dim NetIncome As String
    
    NetIncome = "Net Income"
    
    On Error GoTo ErrorHandler
    
    'Net Income
    Range("A:A").Select
    Selection.Find(What:=NetIncome, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    NetIncome1 = Selection.Offset(0, 1).value
    NetIncome2 = Selection.Offset(0, 2).value
    NetIncome3 = Selection.Offset(0, 3).value
    NetIncome4 = Selection.Offset(0, 4).value
    NetIncome5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Net Income information."
    
    NetIncome1 = 0
    NetIncome2 = 0
    NetIncome3 = 0
    NetIncome4 = 0
    NetIncome5 = 0
    
End Sub

Sub GetEPS()

    Dim EPS As String

    EPS = "Diluted EPS"
    
    On Error GoTo ErrorHandler
    
    'Dividend Per Share
    Columns("A:A").Select
    Selection.Find(What:=EPS, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    EPS1 = Selection.Offset(0, 1).value
    EPS2 = Selection.Offset(0, 2).value
    EPS3 = Selection.Offset(0, 3).value
    EPS4 = Selection.Offset(0, 4).value
    EPS5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No EPS information."
    
    EPS1 = 0
    EPS2 = 0
    EPS3 = 0
    EPS4 = 0
    EPS5 = 0

End Sub

Sub GetDividendPerShare()

    Dim DividendPerShare As String

    DividendPerShare = "Dividend Per Share"
    
    On Error GoTo ErrorHandler
    
    'Diluted EPS
    Columns("A:A").Select
    Selection.Find(What:=DividendPerShare, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    DividendPerShare1 = Selection.Offset(0, 1).value
    DividendPerShare2 = Selection.Offset(0, 2).value
    DividendPerShare3 = Selection.Offset(0, 3).value
    DividendPerShare4 = Selection.Offset(0, 4).value
    DividendPerShare5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Dividend Per Share information."
    
    DividendPerShare1 = 0
    DividendPerShare2 = 0
    DividendPerShare3 = 0
    DividendPerShare4 = 0
    DividendPerShare5 = 0

End Sub
