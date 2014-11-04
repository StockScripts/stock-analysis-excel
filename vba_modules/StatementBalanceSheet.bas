Attribute VB_Name = "StatementBalanceSheet"
Option Explicit

Global dblReceivables(0 To 4) As Double
Global dblInventory(0 To 4) As Double
Global dblCurrentAssets(0 To 4) As Double
Global dblAssets(0 To 4) As Double
Global dblCurrentLiabilities(0 To 4) As Double
Global dblTotalDebt(0 To 4) As Double
Global dblEquity(0 To 4) As Double
Global dblLiabilities(0 To 4) As Double

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
' Notes:        using html DOM screen scraping
'
'               Google Financial statement page source:
'               <div id="balannualdiv" class="id-balannualdiv" style="display:none">            --> parent element
'                   <div id="balannualdiv_viz" class="id-balannualdiv_viz viz_charts"></div>    --> child(0)
'                   <table id=fs-table class="gf-table rgt">                                    --> child(1)/parent
'                       <thead>                                                                 --> child(0)
'                           <tr>                                                                --> children
'                           :
'                       <tbody>                                                                 --> child(1)/parent
'                       <!-- 1 row for each account item -->
'                           <tr>                                                                --> children
'                           <tr>
'                           :
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  09Sept2014 by Janice Laset Parkerson
'               - Initial Version
'
'               07Oct2014 by Janice Laset Parkerson
'               - msn money changed financial data format
'               - no longer able to read via Excel Data Web Query
'               - read financials from Google Finance by screen scraping
'===============================================================
Sub GetAnnualDataBalanceSheet()

    Dim elBalSheetStatement As IHTMLElement
    Dim elColBalSheetStatement As IHTMLElementCollection
    
    Dim elBalSheetTable As IHTMLElement
    Dim elColBalSheetTable As IHTMLElementCollection
        
    Dim elBalSheetDate As IHTMLElement
    Dim elColBalSheetDate As IHTMLElementCollection
    
    Dim elAccountItemData As IHTMLElement
    Dim elColAccountItemData As IHTMLElementCollection
    
    Dim elData As IHTMLElement
    Dim elColData As IHTMLElementCollection

    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    
    'find annual balance sheet statemement <div id="balannualdiv">
    Set elBalSheetStatement = htmlFinancialStatement.getElementById("balannualdiv")
    'get child elements of elBalSheetStatement
    Set elColBalSheetStatement = elBalSheetStatement.Children
    
    'child element(1) of elBalSheetStatement is data table <table id=fs-table>
    Set elBalSheetTable = elColBalSheetStatement(1)
    'get child elements of elBalSheetTable
    Set elColBalSheetTable = elBalSheetTable.Children
    
    'get date information
    'child element(0) of elBalSheetTable is head of data table <thead> (years)
    Set elBalSheetDate = elColBalSheetTable(0)
    'get child element of elBalSheetDate
    Set elColBalSheetDate = elBalSheetDate.Children
    
    'child element(0) of elBalSheetDate is date information
    Set elColData = elColBalSheetDate(0).Children
    For i = 1 To 4
        ActiveSheet.Range("A1").Offset(0, i).Value = elColData(i).innerText
        
        If Err = ERROR_CODE_OBJ_VAR_NOT_SET Then
            ActiveSheet.Range("A1").Offset(0, i).Value = Null
        End If
    Next i
    
    'get balance sheet information
    'child element(1) of elBalSheetTable is body of data table <tbody>
    Set elAccountItemData = elColBalSheetTable(1)
    'get child elements of elAccountItemData (balance sheet items)
    Set elColAccountItemData = elAccountItemData.Children
    
    j = 0
    'get balance sheet items
    For Each elData In elColAccountItemData
        'get child elements for each item (data per year)
        Set elColData = elData.Children
        
        'child (0) is row information
        ActiveSheet.Range("A2").Offset(j, 0).Value = elColData(0).innerText
        
        'children (1 to 4) are row data
        For i = 1 To 4
            ActiveSheet.Range("A2").Offset(j, i).Value = elColData(i).innerText
        Next i
        j = j + 1
    Next elData
    
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
    
    GetReceivables
    GetInventory
    GetCurrentAssets
    GetTotalAssets
    GetCurrentLiabilities
    GetTotalDebt
    GetLiabilities
    GetEquity
    
    Columns("A:E").EntireColumn.AutoFit
    
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
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    Receivables = "Total Receivables, Net "
    
    'find receivables account item
    Columns("A:A").Select
    Selection.Find(What:=Receivables, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblReceivables(i) = Selection.Offset(0, i + 1).Value
    Next i
        
    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Receivables information."
   
    For i = 0 To 3
        dblReceivables(i) = 0
    Next i

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
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    Inventory = "Total Inventory "
    
    'find inventory account item
    Columns("A:A").Select
    Selection.Find(What:=Inventory, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblInventory(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
        
    Exit Sub
        
ErrorHandler:
   MsgBox "No Inventory information."
   
    For i = 0 To 3
        dblInventory(i) = 0
    Next i
   
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
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    CurrentAssets = "Total Current Assets "
        
    'find current assets account item
    Columns("A:A").Select
    Selection.Find(What:=CurrentAssets, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblCurrentAssets(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Current Assets information."
    
    For i = 0 To 3
        dblCurrentAssets(i) = 0
    Next i

End Sub

'===============================================================
' Procedure:    GetTotalAssets
'
' Description:  Find total assets information in balance sheet
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
' Rev History:   11Oct2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetTotalAssets()

    Dim strTotalAssets As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    strTotalAssets = "Total Assets "
        
    'find current assets account item
    Columns("A:A").Select
    Selection.Find(What:=strTotalAssets, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblAssets(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Total Assets information."
    
    For i = 0 To 3
        dblAssets(i) = 0
    Next i

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
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    CurrentLiabilities = "Total Current Liabilities "
        
    'find current liabilities account item
    Columns("A:A").Select
    Selection.Find(What:=CurrentLiabilities, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblCurrentLiabilities(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font

    Exit Sub
    
ErrorHandler:
    MsgBox "No Current Liabilities information."
    
    For i = 0 To 3
        dblCurrentLiabilities(i) = 0
    Next i
    
End Sub

'===============================================================
' Procedure:    GetTotalDebt
'
' Description:  Find total debt information in balance sheet
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
Sub GetTotalDebt()

    Dim TotalDebt As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    TotalDebt = "Total Debt "
    
    'find long term debt account item
    Columns("A:A").Select
    Selection.Find(What:=TotalDebt, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblTotalDebt(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Total Debt information."
    
    For i = 0 To 3
        dblTotalDebt(i) = 0
    Next i
    
End Sub

'===============================================================
' Procedure:    GetLiabilities
'
' Description:  Find total liabilities information in balance sheet
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
' Rev History:   09Oct2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetLiabilities()

    Dim Liabilities As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    Liabilities = "Total Liabilities "
    
    'find long term debt account item
    Columns("A:A").Select
    Selection.Find(What:=Liabilities, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblLiabilities(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = 5           'blue font
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Liabilities information."
    
    For i = 0 To 3
        dblLiabilities(i) = 0
    Next i
    
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
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    Equity = "Total Equity "
       
    'find equity account item
    Columns("A:A").Select
    Selection.Find(What:=Equity, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblEquity(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Equity information."
    
    For i = 0 To 3
        dblEquity(i) = 0
    Next i
    
End Sub
