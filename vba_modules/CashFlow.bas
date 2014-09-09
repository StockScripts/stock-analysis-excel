Attribute VB_Name = "CashFlow"
Option Explicit

Public OpCashFlow1 As Double
Public OpCashFlow2 As Double
Public OpCashFlow3 As Double
Public OpCashFlow4 As Double
Public OpCashFlow5 As Double

Public FreeCashFlow1 As Double
Public FreeCashFlow2 As Double
Public FreeCashFlow3 As Double
Public FreeCashFlow4 As Double
Public FreeCashFlow5 As Double

'Create Cash Flow statement with data from msnmoney.com
Sub CashFlowStatement()

    CreateCashFlowSheet
    GetCashFlow
    FormatCashFlow

End Sub

Sub CreateCashFlowSheet()

    Dim oSheet As Worksheet, vRet As Variant

    On Error GoTo ErrorHandler
    
    Set oSheet = Worksheets.Add
    With oSheet
        .Name = "Cash Flow - " & TickerSym
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
            Worksheets("Cash Flow - " & TickerSym).Delete
            Application.DisplayAlerts = True

            'rename and activate the new worksheet
            With oSheet
                .Name = "Cash Flow - " & TickerSym
                .Cells(1.1).Select
                .Activate
            End With
        Else
            'cancel the operation, delete the new worksheet
            Application.DisplayAlerts = False
            oSheet.Delete
            Application.DisplayAlerts = True
            'activate the old worksheet
            Worksheets("Cash Flow - " & TickerSym).Activate
        End If

    End If
    
End Sub

'Gets data from msnmoney.com
Sub GetCashFlow()

    On Error GoTo ErrorHandler
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://investing.money.msn.com/investments/stock-cash-flow/?symbol=" & TickerSym & "", _
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

'Assign account names to cells in Cash Flow Statement
Sub FormatCashFlow()
    
    Sheets("Cash Flow - " & TickerSym).Activate
    
    GetOpCashFlow
    GetFreeCashFlow
        
End Sub

Sub GetOpCashFlow()

    Dim OpCashFlow As String
    
    OpCashFlow = "Cash Flow from Operating Activities"
    
    On Error GoTo ErrorHandler
    
    'Cash from Operating Activities
    Columns("A:A").Select
    Selection.Find(What:=OpCashFlow, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    OpCashFlow1 = Selection.Offset(0, 1).value
    OpCashFlow2 = Selection.Offset(0, 2).value
    OpCashFlow3 = Selection.Offset(0, 3).value
    OpCashFlow4 = Selection.Offset(0, 4).value
    OpCashFlow5 = Selection.Offset(0, 5).value

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
   MsgBox "No Operating Cash Flow information."
   
   OpCashFlow1 = 0
   OpCashFlow2 = 0
   OpCashFlow3 = 0
   OpCashFlow4 = 0
   OpCashFlow5 = 0
    
End Sub

Sub GetFreeCashFlow()

    Dim FreeCashFlow As String

    FreeCashFlow = "Free Cash Flow"
    
    On Error GoTo ErrorHandler
    
    Columns("A:A").Select
    Selection.Find(What:=FreeCashFlow, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    FreeCashFlow1 = Selection.Offset(0, 1).value
    FreeCashFlow2 = Selection.Offset(0, 2).value
    FreeCashFlow3 = Selection.Offset(0, 3).value
    FreeCashFlow4 = Selection.Offset(0, 4).value
    FreeCashFlow5 = Selection.Offset(0, 5).value
    
    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
   MsgBox "No Free Cash Flow information."
    
   FreeCashFlow1 = 0
   FreeCashFlow2 = 0
   FreeCashFlow3 = 0
   FreeCashFlow4 = 0
   FreeCashFlow5 = 0
End Sub
