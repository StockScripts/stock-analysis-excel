Attribute VB_Name = "Checklist"
Option Explicit

Sub StocksChecklist()

    CreateCheckList
    FormatCheckList
    
    Revenue
    Profit
    QuickRatio
    Debt
    FreeCashFlow
    EPS
    RedFlags
    DividendPerShare
    PriceToEarnings
    
End Sub

Sub CreateCheckList()

    Dim oSheet As Worksheet, vRet As Variant

    On Error GoTo ErrorHandler
    
    Set oSheet = Worksheets.Add
    With oSheet
        .Name = "Analysis - " & TickerSym
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
            Worksheets("Analysis - " & TickerSym).Delete
            Application.DisplayAlerts = True

            'rename and activate the new worksheet
            With oSheet
                .Name = "Analysis - " & TickerSym
                .Cells(1.1).Select
                .Activate
            End With
        Else
            'cancel the operation, delete the new worksheet
            Application.DisplayAlerts = False
            oSheet.Delete
            Application.DisplayAlerts = True
            'activate the old worksheet
            Worksheets("Analysis - " & TickerSym).Activate
        End If

    End If

End Sub

Sub FormatCheckList()

    Worksheets("Analysis - " & TickerSym).Activate
    Range("1:1").Font.Bold = True
    Range("1:1").HorizontalAlignment = xlCenter
    
    Range("A:A").ColumnWidth = 5
    Range("B:B").ColumnWidth = 19
    Range("C:C").ColumnWidth = 9
    Range("D:D").ColumnWidth = 9
    Range("E:E").ColumnWidth = 9
    Range("F:F").ColumnWidth = 9
    Range("G:G").ColumnWidth = 9
    
    Range("C1") = Sheets("Balance Sheet - " & TickerSym).Range("Year1")
    Range("D1") = Sheets("Balance Sheet - " & TickerSym).Range("Year2")
    Range("E1") = Sheets("Balance Sheet - " & TickerSym).Range("Year3")
    Range("F1") = Sheets("Balance Sheet - " & TickerSym).Range("Year4")
    Range("G1") = Sheets("Balance Sheet - " & TickerSym).Range("Year5")
    
    Range("A33").Font.Bold = True
    Range("A33") = "Can they pay back investors?"
    
    Range("A36").Font.Bold = True
    Range("A36") = "Is it overpriced?"
    
End Sub

