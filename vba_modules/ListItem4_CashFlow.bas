Attribute VB_Name = "ListItem4_CashFlow"
Option Explicit

Private ResultCashFlow As Result

'===============================================================
' Procedure:    EvaluateFreeCashFlow
'
' Description:  Display free cash flow information.
'               Only Cash flow for most recent year must be positive.
'               if recent year cash flow > 0 -> green font -> pass
'               else -> red font -> fail
'               if past years cash flow > 0 -> green font -> pass
'               else -> orange font -> pass
'               Call procedure to display YOY growth information
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  10Octt14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateFreeCashFlow()

    Dim i As Integer
    
    Range("ListItemFreeCashFlow") = "Is there free cash flow?"
    Range("FreeCashFlow") = "Free Cash Flow"
    
    ResultCashFlow = PASS
    
'   populate free cash flow information
    If dblFreeCashFlow(0) > 0 Then
        Range("FreeCashFlow").Offset(0, 1).Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("FreeCashFlow").Offset(0, 1).Font.ColorIndex = FONT_COLOR_RED
        ResultCashFlow = FAIL
    End If
    Range("FreeCashFlow").Offset(0, 1) = dblFreeCashFlow(0)
    
    For i = 1 To (iYearsAvailableIncome - 1)
        If dblFreeCashFlow(i) > 0 Then
            Range("FreeCashFlow").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_GREEN
        Else
            Range("FreeCashFlow").Offset(0, i + 1).Font.ColorIndex = FONT_COLOR_ORANGE
        End If
        Range("FreeCashFlow").Offset(0, i + 1) = dblFreeCashFlow(i)
    Next i
    
    DisplayFreeCashFlowInfo
    
    CalculateFreeCashFlowYOYGrowth

End Sub

'===============================================================
' Procedure:    DisplayEarningsInfo
'
' Description:  Comment box information for free cash flow
'               - cash flow requirements
'               - cash flow formula
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  10Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub DisplayFreeCashFlowInfo()

    With Range("ListItemFreeCashFlow")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="What is it:" & Chr(10) & _
                "   Free cash flow is cash that a company generates after paying expenses." & Chr(10) & _
                "Why is it important:" & Chr(10) & _
                "   Free cash flow enhances value by allowing a company to develop new products, make" & Chr(10) & _
                "   acquisitions, pay dividends, or reduce debt." & Chr(10) & _
                "What to look for:" & Chr(10) & _
                "   The recent year free cash flow should be positive." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   Free cash flow should not be continuously decreasing."
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    With Range("FreeCashFlow")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="Free Cash flow = Operating Cash Flow - Capital Expenditures"
        .Comment.Shape.TextFrame.AutoSize = True
    End With

End Sub

'===============================================================
' Procedure:    CalculateFreeCashFlowYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for free cash flow data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  10Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateFreeCashFlowYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double
    Dim i As Integer

    Range("FreeCashFlowYOYGrowth") = "YOY Growth (%)"
    
    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To (iYearsAvailableIncome - 2)
        dblYOYGrowth(i) = CalculateYOYGrowth(dblFreeCashFlow(i), dblFreeCashFlow(i + 1))
    Next i
    
    Call EvaluateFreeCashFlowYOYGrowth(Range("FreeCashFlowYOYGrowth"), dblYOYGrowth)
    
End Sub

'===============================================================
' Procedure:    EvaluateFreeCashFlowYOYGrowth
'
' Description:  Display YOY growth information.
'               if recent year free cash flow <= 0 -> red font -> fail
'               else if free cash flow decreases -> orange font -> pass
'               else green font -> pass
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   YOYGrowth As Range -> first cell of revenue YOY growth
'               YOY1, YOY2, YOY3 -> YOY growth values
'                                   (YOY1 is most recent year)
'
' Returns:      N/A
'
' Rev History:  10Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateFreeCashFlowYOYGrowth(YOYGrowth As Range, YOY() As Double)
    
    Dim i As Integer
    
    For i = 0 To (iYearsAvailableIncome - 2)
        YOYGrowth.Offset(0, i + 1).Select
        If dblFreeCashFlow(i) <= 0 Then
            Selection.Font.ColorIndex = FONT_COLOR_RED
        ElseIf YOY(i) < 0 Then
            Selection.Font.ColorIndex = FONT_COLOR_ORANGE
        Else
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
        End If
        YOYGrowth.Offset(0, i + 1) = YOY(i)
    Next i
    
    CheckCashFlowPassFail
    
End Function

'===============================================================
' Procedure:    CheckCashFlowPassFail
'
' Description:  Display check or x mark if the cash flow
'               passes or fails the criteria
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  10Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CheckCashFlowPassFail()

    If ResultCashFlow = PASS Then
        Range("FreeCashflowCheck") = CHECK_MARK
        Range("FreeCashflowCheck").Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("FreeCashflowCheck") = X_MARK
        Range("FreeCashflowCheck").Font.ColorIndex = FONT_COLOR_RED
    End If

End Sub
