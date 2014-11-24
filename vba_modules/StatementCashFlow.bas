Attribute VB_Name = "StatementCashFlow"
Option Explicit

Global dblOpCashFlow(0 To 4) As Double
Global dblCapEx(0 To 4) As Double
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
' Notes:        using html DOM screen scraping
'
'               Google Financial statement page source:
'               <div id="casannualdiv" class="id-casannualdiv" style="display:none">            --> parent element
'                   <div id="casannualdiv_viz" class="id-casannualdiv_viz viz_charts"></div>    --> child(0)
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
' Rev History:  11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'
'               07Oct2014 by Janice Laset Parkerson
'               - msn money changed financial data format
'               - no longer able to read via Excel Data Web Query
'               - read financials from Google Finance by screen scraping
'===============================================================
Sub GetAnnualDataCashFlow()

    Dim elCashFlowStatement As IHTMLElement
    Dim elColCashFlowStatement As IHTMLElementCollection
    
    Dim elCashFlowTable As IHTMLElement
    Dim elColCashFlowTable As IHTMLElementCollection
        
    Dim elCashFlowDate As IHTMLElement
    Dim elColCashFlowDate As IHTMLElementCollection
    
    Dim elAccountItemData As IHTMLElement
    Dim elColAccountItemData As IHTMLElementCollection
    
    Dim elData As IHTMLElement
    Dim elColData As IHTMLElementCollection

    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    
    'find annual cash flow statemement <div id="casannualdiv">
    Set elCashFlowStatement = htmlFinancialStatement.getElementById("casannualdiv")
    'get child elements of elCashFlowStatement
    Set elColCashFlowStatement = elCashFlowStatement.Children
    
    'child element(1) of elCashFlowStatement is data table <table id=fs-table>
    Set elCashFlowTable = elColCashFlowStatement(1)
    'get child elements of elCashFlowTable
    Set elColCashFlowTable = elCashFlowTable.Children
         
    'get date information
    'child element(0) of elCashFlowTable is head of data table <thead> (years)
    Set elCashFlowDate = elColCashFlowTable(0)
    'get child element of elCashFlowDate
    Set elColCashFlowDate = elCashFlowDate.Children
    
    'child element(0) of elCashFlowDate is date information
    Set elColData = elColCashFlowDate(0).Children
    For i = 1 To 4
        ActiveSheet.Range("A1").Offset(0, i).Value = elColData(i).innerText
        
        If Err Then
            ActiveSheet.Range("A1").Offset(0, i).Value = Null
        End If
    Next i
    
    'get cash flow information
    'child element(1) of elCashFlowTable is body of data table <tbody>
    Set elAccountItemData = elColCashFlowTable(1)
    'get child elements of elAccountItemData (cash flow statement items)
    Set elColAccountItemData = elAccountItemData.Children
    
    j = 0
    'get cash flow statement items
    For Each elData In elColAccountItemData
        'get child elements (cash flow statement items)
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
    GetCapEx
    GetFreeCashFlow
    
    Columns("A:E").EntireColumn.AutoFit
        
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
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    OpCashFlow = "Cash from Operating Activities "
    
    'find operating cash flow account item
    Columns("A:A").Select
    Selection.Find(What:=OpCashFlow, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblOpCashFlow(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
   MsgBox "No Operating Cash Flow information."
   
    For i = 0 To 3
        dblOpCashFlow(i) = 0
    Next i
    
End Sub

'===============================================================
' Procedure:    GetCapEx
'
' Description:  Find capital expenditures information in cash flow
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
' Rev History:   07Oct2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetCapEx()

    Dim CapEx As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in balance sheet
    CapEx = "Capital Expenditures "
    
    'find operating cash flow account item
    Columns("A:A").Select
    Selection.Find(What:=CapEx, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblCapEx(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
   MsgBox "No CapitalExpenditures information."
   
    For i = 0 To 3
        dblCapEx(i) = 0
    Next i

End Sub

'===============================================================
' Procedure:    GetFreeCashFlow
'
' Description:  Calculate free cash flow.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  11Sept2014 by Janice Laset Parkerson
'               - Initial Version
'
'               07Oct2014 by Janice Laset Parkerson
'               - calculate free cash flow (not available on google
'                 finance page)
'===============================================================
Sub GetFreeCashFlow()

    On Error GoTo ErrorHandler
    Dim i As Integer
        
    'free cash flow = operating cash flow - capital expenditures
    'cap ex recorded as negative value in statement -> add to op cash flow
    For i = 0 To 3
        dblFreeCashFlow(i) = dblOpCashFlow(i) + dblCapEx(i)
    Next i
    
    Exit Sub
    
ErrorHandler:
   MsgBox "No Free Cash Flow information."
    
    For i = 0 To 3
        dblFreeCashFlow(i) = 0
    Next i

End Sub
