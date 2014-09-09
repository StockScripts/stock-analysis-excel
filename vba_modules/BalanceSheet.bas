Attribute VB_Name = "BalanceSheet"
Option Explicit

Public Receivables1 As Double
Public Receivables2 As Double
Public Receivables3 As Double
Public Receivables4 As Double
Public Receivables5 As Double

Public Inventory1 As Double
Public Inventory2 As Double
Public Inventory3 As Double
Public Inventory4 As Double
Public Inventory5 As Double

Public CurrentAssets1 As Double
Public CurrentAssets2 As Double
Public CurrentAssets3 As Double
Public CurrentAssets4 As Double
Public CurrentAssets5 As Double

Public CurrentLiabilities1 As Double
Public CurrentLiabilities2 As Double
Public CurrentLiabilities3 As Double
Public CurrentLiabilities4 As Double
Public CurrentLiabilities5 As Double

Public LongTermDebt1 As Double
Public LongTermDebt2 As Double
Public LongTermDebt3 As Double
Public LongTermDebt4 As Double
Public LongTermDebt5 As Double

Public Equity1 As Double
Public Equity2 As Double
Public Equity3 As Double
Public Equity4 As Double
Public Equity5 As Double

'Create Cash Flow Worksheet with data from msnmoney.com
Sub BalanceSheetStatement()

    CreateBalanceSheet
    GetBalanceSheet
    FormatBalanceSheet

End Sub

Sub CreateBalanceSheet()

    Dim oSheet As Worksheet, vRet As Variant

    On Error GoTo ErrorHandler
    
    Set oSheet = Worksheets.Add
    With oSheet
        .Name = "Balance Sheet - " & TickerSym
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
            Worksheets("Balance Sheet - " & TickerSym).Delete
            Application.DisplayAlerts = True

            'rename and activate the new worksheet
            With oSheet
                .Name = "Balance Sheet - " & TickerSym
                .Cells(1.1).Select
                .Activate
            End With
        Else
            'cancel the operation, delete the new worksheet
            Application.DisplayAlerts = False
            oSheet.Delete
            Application.DisplayAlerts = True
            'activate the old worksheet
            Worksheets("Balance Sheet - " & TickerSym).Activate
        End If

    End If
    
End Sub


'Gets annual Balance Sheet statement from msnmoney.com
Sub GetBalanceSheet()

    On Error GoTo ErrorHandler
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://investing.money.msn.com/investments/stock-balance-sheet/?symbol=us%3A" & TickerSym & "&stmtView=Ann" _
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

'Assign account names to cells in Balance Sheet
Sub FormatBalanceSheet()

    Sheets("Balance Sheet - " & TickerSym).Activate
    
    GetReceivables
    GetInventory
    GetCurrentAssets
    GetCurrentLiabilities
    GetLongTermDebt
    GetEquity
    
    'Get Years
    Sheets("Balance Sheet - " & TickerSym).Range("B1").Name = "Year1"
    Sheets("Balance Sheet - " & TickerSym).Range("C1").Name = "Year2"
    Sheets("Balance Sheet - " & TickerSym).Range("D1").Name = "Year3"
    Sheets("Balance Sheet - " & TickerSym).Range("E1").Name = "Year4"
    Sheets("Balance Sheet - " & TickerSym).Range("F1").Name = "Year5"

End Sub

Sub GetReceivables()

    Dim Receivables As String

    Receivables = "Receivables"
    
    On Error GoTo ErrorHandler
    
    'Receivables
    Columns("A:A").Select
    Selection.Find(What:=Receivables, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    Receivables1 = Selection.Offset(0, 1).value
    Receivables2 = Selection.Offset(0, 2).value
    Receivables3 = Selection.Offset(0, 3).value
    Receivables4 = Selection.Offset(0, 4).value
    Receivables5 = Selection.Offset(0, 5).value
        
    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Receivables information."
   
    Receivables1 = 0
    Receivables2 = 0
    Receivables3 = 0
    Receivables4 = 0
    Receivables5 = 0

End Sub

Sub GetInventory()

    Dim Inventory As String
    
    Inventory = "Inventories"
    
    On Error GoTo ErrorHandler
    
    Columns("A:A").Select
    Selection.Find(What:=Inventory, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    Inventory1 = Selection.Offset(0, 1).value
    Inventory2 = Selection.Offset(0, 2).value
    Inventory3 = Selection.Offset(0, 3).value
    Inventory4 = Selection.Offset(0, 4).value
    Inventory5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
        
    Exit Sub
        
ErrorHandler:
   MsgBox "No Inventory information."
   
   Inventory1 = 0
   Inventory2 = 0
   Inventory3 = 0
   Inventory4 = 0
   Inventory5 = 0
   
End Sub

Sub GetCurrentAssets()

    Dim CurrentAssets As String
    
    CurrentAssets = "Total Current Assets"
    
    On Error GoTo ErrorHandler
    
    'Current Assets
    Columns("A:A").Select
    Selection.Find(What:=CurrentAssets, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    CurrentAssets1 = Selection.Offset(0, 1).value
    CurrentAssets2 = Selection.Offset(0, 2).value
    CurrentAssets3 = Selection.Offset(0, 3).value
    CurrentAssets4 = Selection.Offset(0, 4).value
    CurrentAssets5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Current Assets information."
    
    CurrentAssets1 = 0
    CurrentAssets2 = 0
    CurrentAssets3 = 0
    CurrentAssets4 = 0
    CurrentAssets5 = 0

End Sub

Sub GetCurrentLiabilities()

    Dim CurrentLiabilities As String
    
    CurrentLiabilities = "Total Current Liabilities"
    
    On Error GoTo ErrorHandler
    
    'Current Liabilities
    Columns("A:A").Select
    Selection.Find(What:=CurrentLiabilities, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    CurrentLiabilities1 = Selection.Offset(0, 1).value
    CurrentLiabilities2 = Selection.Offset(0, 2).value
    CurrentLiabilities3 = Selection.Offset(0, 3).value
    CurrentLiabilities4 = Selection.Offset(0, 4).value
    CurrentLiabilities5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font

    Exit Sub
    
ErrorHandler:
    MsgBox "No Current Liabilities information."
    
    CurrentLiabilities1 = 0
    CurrentLiabilities2 = 0
    CurrentLiabilities3 = 0
    CurrentLiabilities4 = 0
    CurrentLiabilities5 = 0
    
End Sub

Sub GetLongTermDebt()

    Dim LongTermDebt As String
    
    LongTermDebt = "Lt Debt and Capital Lease Obligation"
    
    On Error GoTo ErrorHandler
    
    'Long Term Debt
    Columns("A:A").Select
    Selection.Find(What:=LongTermDebt, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    LongTermDebt1 = Selection.Offset(0, 1).value
    LongTermDebt2 = Selection.Offset(0, 2).value
    LongTermDebt3 = Selection.Offset(0, 3).value
    LongTermDebt4 = Selection.Offset(0, 4).value
    LongTermDebt5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Long Term Debt information."
    
    LongTermDebt1 = 0
    LongTermDebt2 = 0
    LongTermDebt3 = 0
    LongTermDebt4 = 0
    LongTermDebt5 = 0
    
End Sub

Sub GetEquity()

    Dim Equity As String
    
    Equity = "Total Equity"
    
    On Error GoTo ErrorHandler
    
    'Equity
    Columns("A:A").Select
    Selection.Find(What:=Equity, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    Equity1 = Selection.Offset(0, 1).value
    Equity2 = Selection.Offset(0, 2).value
    Equity3 = Selection.Offset(0, 3).value
    Equity4 = Selection.Offset(0, 4).value
    Equity5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Equity information."
    
    Equity1 = 0
    Equity2 = 0
    Equity3 = 0
    Equity4 = 0
    Equity5 = 0
    
End Sub
