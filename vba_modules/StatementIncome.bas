Attribute VB_Name = "StatementIncome"
Option Explicit

' variables may be string '---' if no data
Global dblRevenue(0 To 4) As Double
Global dblSGA(0 To 4) As Double
Global dblIncomeBeforeTax(0 To 4) As Double
Global dblIncomeAfterTax(0 To 4) As Double
Global dblNetIncome(0 To 4) As Double
Global dblOperatingExpense(0 To 4) As Double
Global dblShares(0 To 4) As Double
Global vEPS(0 To 4) As Variant
Global dblDividendPerShare(0 To 4) As Double
Global strYear(0 To 4) As String
Global iYearsAvailableIncome As Integer

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
' Description:  Get annual income statement from Google Finance
'
' Author:       Janice Laset Parkerson
'
' Notes:        using html DOM screen scraping
'
'               Google Financial statement page source:
'               <div id="incannualdiv" class="id-incannualdiv" style="display:none">            --> parent element
'                   <div id="incannualdiv_viz" class="id-incannualdiv_viz viz_charts"></div>    --> child(0)
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
Sub GetAnnualDataIncome()

    Dim elIncStatement As IHTMLElement
    Dim elColIncStatement As IHTMLElementCollection
    
    Dim elIncomeTable As IHTMLElement
    Dim elColIncomeTable As IHTMLElementCollection
    
    Dim elIncomeDate As IHTMLElement
    Dim elColIncomeDate As IHTMLElementCollection
        
    Dim elAccountItemData As IHTMLElement
    Dim elColAccountItemData As IHTMLElementCollection
    
    Dim elData As IHTMLElement
    Dim elColData As IHTMLElementCollection

    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    
    'find annual income statemement <div id="incannualdiv">
    Set elIncStatement = htmlFinancialStatement.getElementById("incannualdiv")
    'get child elements of elIncStatement
    Set elColIncStatement = elIncStatement.Children
    
    'child element(1) of elIncStatement is data table <table id=fs-table>
    Set elIncomeTable = elColIncStatement(1)
    'get child elements of elIncomeTable
    Set elColIncomeTable = elIncomeTable.Children
    
    'get date information
    'child element(0) of elIncomeTable is head of data table <thead> (years)
    Set elIncomeDate = elColIncomeTable(0)
    'get child element of elIncomeDate
    Set elColIncomeDate = elIncomeDate.Children
    
    'child element(0) of elIncomeDate is date information
    Set elColData = elColIncomeDate(0).Children
    For i = 1 To YEARS_MAX
        ActiveSheet.Range("A1").Offset(0, i).Value = elColData(i).innerText
         
        'if statement has less than 4 years of data
        If Err = ERROR_CODE_OBJ_VAR_NOT_SET Then
            ActiveSheet.Range("A1").Offset(0, i).Value = vbNullString
            iYearsAvailableIncome = i - 1   'get max years of available income data
            Err.Clear
            Exit For
        End If
        
        iYearsAvailableIncome = YEARS_MAX
    Next i
    
    'get income statement information
    'child element(1) of elIncomeTable is body of data table <tbody>
    Set elAccountItemData = elColIncomeTable(1)
    'get child elements of elAccountItemData (income statement items)
    Set elColAccountItemData = elAccountItemData.Children
    
    j = 0
    'get income statement items
    For Each elData In elColAccountItemData
        'get child elements for each item (data per year)
        Set elColData = elData.Children
        
        'child(0) is row information
        ActiveSheet.Range("A2").Offset(j, 0).Value = elColData(0).innerText
        
        'children (1 to 4) are row data
        For i = 1 To iYearsAvailableIncome
            ActiveSheet.Range("A2").Offset(j, i).Value = elColData(i).innerText
        Next i
        j = j + 1
    Next elData
    
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
    
    GetYears
    GetRevenue
    GetSGA
    GetOperatingExpense
    GetIncomeBeforeTax
    GetIncomeAfterTax
    GetNetIncome
    GetShares
    GetEPS
    GetDividendPerShare
    
    Columns("A:E").EntireColumn.AutoFit
    
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

    Dim strReportDate(0 To 3) As String
    Dim i As Integer
    'year header text is "12 months ending YYYY-MM-DD "
    'use Mid string function to extract date
    'YYYY-MM-DD begins at index 1 (base 1)
    Const iYearIndex As Integer = 17
        
    For i = 0 To (iYearsAvailableIncome - 1)
        strReportDate(i) = Range("A1").Offset(0, i + 1).Value
        strYear(i) = Mid(strReportDate(i), iYearIndex)
    Next i

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
    Dim i As Integer
       
    'account item term to search for in income statement
    Revenue = "Total Revenue "
    
    'find revenue account item
    Columns("A:A").Select
    Selection.Find(What:=Revenue, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select

    For i = 0 To (iYearsAvailableIncome - 1)
        dblRevenue(i) = Selection.Offset(0, i + 1).Value
    Next i
    
    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
        
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
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in income statement
    SGA = "Selling/General/Admin. Expenses, Total "
       
    'find SGA account item
    Columns("A:A").Select
    Selection.Find(What:=SGA, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblSGA(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No SGA information."
    
    For i = 0 To 3
        dblSGA(i) = 0
    Next i
    
End Sub

'===============================================================
' Procedure:    GetOperatingExpense
'
' Description:  Find operating expense information in income statement
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
' Rev History:   27Oct2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetOperatingExpense()

    Dim strOperatingExpense As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'account item term to search for in income statement
    strOperatingExpense = "Total Operating Expense "
       
    'find SGA account item
    Columns("A:A").Select
    Selection.Find(What:=strOperatingExpense, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblOperatingExpense(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Operating Expense information."
    
    For i = 0 To 3
        dblOperatingExpense(i) = 0
    Next i
    
End Sub

'===============================================================
' Procedure:    GetIncomeBeforeTax
'
' Description:  Find income before tax information in income statement
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
' Rev History:   28Oct2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetIncomeBeforeTax()

    Dim strIncomeBeforeTax As String
    Dim i As Integer
    
    On Error Resume Next
    
    'account item term to search for in income statement
    strIncomeBeforeTax = "Income Before Tax "
         
    'find net income account item
    Range("A:A").Select
    Selection.Find(What:=strIncomeBeforeTax, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To (iYearsAvailableIncome - 1)
        dblIncomeBeforeTax(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE


End Sub

'===============================================================
' Procedure:    GetIncomeAfterTax
'
' Description:  Find income after tax information in income statement
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
' Rev History:   28Oct2014 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetIncomeAfterTax()

    Dim strIncomeAfterTax As String
    Dim i As Integer
    
    On Error Resume Next
    
    'account item term to search for in income statement
    strIncomeAfterTax = "Income After Tax "
         
    'find net income account item
    Range("A:A").Select
    Selection.Find(What:=strIncomeAfterTax, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To (iYearsAvailableIncome - 1)
        dblIncomeAfterTax(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE

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
    Dim i As Integer
    
    On Error Resume Next
    
    'account item term to search for in income statement
    NetIncome = "Net Income "
         
    'find net income account item
    Range("A:A").Select
    Selection.Find(What:=NetIncome, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To (iYearsAvailableIncome - 1)
        dblNetIncome(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
End Sub

'===============================================================
' Procedure:    GetShares
'
' Description:  Find shares information in income statement
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
Sub GetShares()

    Dim strShares As String
    Dim i As Integer
    
    On Error Resume Next
    
    'account item term to search for in income statement
    strShares = "Diluted Weighted Average Shares "
         
    'find net income account item
    Range("A:A").Select
    Selection.Find(What:=strShares, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To (iYearsAvailableIncome - 1)
        dblShares(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE

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

    Dim strEPS As String
    Dim i As Integer
    
    'account item term to search for in income statement
    strEPS = "Diluted Normalized EPS "
    
    On Error Resume Next
    
    'find EPS account item
    Columns("A:A").Select
    Selection.Find(What:=strEPS, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To (iYearsAvailableIncome - 1)
        vEPS(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
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
    Dim i As Integer
    
    'account item term to search for in income statement
    DividendPerShare = "Dividends per Share - Common Stock Primary Issue "
    
    On Error GoTo ErrorHandler
    
    'find dividend per share account item
    Columns("A:A").Select
    Selection.Find(What:=DividendPerShare, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Select
        
    For i = 0 To 3
        dblDividendPerShare(i) = Selection.Offset(0, i + 1).Value
    Next i

    Rows(ActiveCell.Row).Select
    Selection.Font.ColorIndex = FONT_COLOR_BLUE
    
    Exit Sub
    
ErrorHandler:
    MsgBox "No Dividend Per Share information."
    
    For i = 0 To 3
        dblDividendPerShare(i) = 0
    Next i

End Sub
