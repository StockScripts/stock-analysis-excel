Attribute VB_Name = "ListItem4_CashFlow"
Option Explicit

Private ResultCashFlow As Result
Private Const CASH_FLOW_SCORE_MAX = 4
Private Const CASH_FLOW_SCORE_WEIGHT = 6
Public ScoreCashFlow As Integer
Public Const MAX_CASH_FLOW_SCORE = 114

'===============================================================
' Procedure:    EvaluateFreeCashFlow
'
' Description:  Display free cash flow information.
'               Cash flow for most recent year must be positive.
'               if recent year cash flow > 0 -> pass
'               else -> fail
'               if past years cash flow > 0 -> pass
'               else -> warning
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
    ScoreCashFlow = 0
    
'   populate free cash flow information
    For i = 1 To (iYearsAvailableIncome - 1)
        Range("FreeCashFlow").Offset(0, i + 1).Select
        If dblFreeCashFlow(i) > 0 Then
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
            ScoreCashFlow = ScoreCashFlow + (CASH_FLOW_SCORE_MAX - i)
        Else
            Selection.Font.ColorIndex = FONT_COLOR_ORANGE
        End If
        Selection.Value = dblFreeCashFlow(i)
    Next i
    
    Range("FreeCashFlow").Offset(0, 1).Select
    If dblFreeCashFlow(0) > 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
        ScoreCashFlow = ScoreCashFlow + CASH_FLOW_SCORE_MAX
    Else
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultCashFlow = FAIL
        ScoreCashFlow = ScoreCashFlow - (CASH_FLOW_SCORE_MAX * 2)
    End If
    Selection.Value = dblFreeCashFlow(0)
    
    DisplayFreeCashFlowInfo
    
    CalculateFreeCashFlowYOYGrowth

End Sub

'===============================================================
' Procedure:    DisplayFreeCashFlowInfo
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

    Dim dblOpCashFlowYOYGrowth(0 To 3) As Double
    Dim strOpCashFlowYOYGrowth(0 To 3) As String
    
    Dim dblCapExYOYGrowth(0 To 2) As Double
    Dim strCapExYOYGrowth(0 To 2) As String
    Dim dblAbsCapEx(0 To 3) As Double
    
    Dim i As Integer
    
    With Range("ListItemFreeCashFlow")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="What is it:" & Chr(10) & _
                "   Free cash flow is cash that a company generates after paying expenses." & Chr(10) & _
                "Why is it important:" & Chr(10) & _
                "   Free cash flow enhances value by allowing a company to develop new products, make" & Chr(10) & _
                "   acquisitions, pay dividends, or reduce debt. Growing free cash flows are frequently" & Chr(10) & _
                "   a prelude to increased earnings." & Chr(10) & _
                "What to look for:" & Chr(10) & _
                "   Free cash flow should ideally be increasing, and the recent year should be positive." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   Free cash flow should not be continuously decreasing."
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    'get absolute value of cap ex - recorded as negative in statement
    For i = 0 To (iYearsAvailableIncome - 1)
        dblAbsCapEx(i) = Abs(dblCapEx(i))
    Next i
    
    'calculate YOY growth
    For i = 0 To (iYearsAvailableIncome - 2)
        dblOpCashFlowYOYGrowth(i) = CalculateYOYGrowth(dblOpCashFlow(i), dblOpCashFlow(i + 1))
        strOpCashFlowYOYGrowth(i) = Format(dblOpCashFlowYOYGrowth(i), "0.0%")
        
        dblCapExYOYGrowth(i) = CalculateYOYGrowth(dblAbsCapEx(i), dblAbsCapEx(i + 1))
        strCapExYOYGrowth(i) = Format(dblCapExYOYGrowth(i), "0.0%")
    Next i
    
    With Range("FreeCashFlow")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="Free Cash flow = Operating Cash Flow - Capital Expenditures" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Operating Cash Flow" & "                " & dblOpCashFlow(0) & "      " & dblOpCashFlow(1) & "      " & dblOpCashFlow(2) & "      " & dblOpCashFlow(3) & Chr(10) & _
                "YOY Operating Cash Flow Growth     " & strOpCashFlowYOYGrowth(0) & "     " & strOpCashFlowYOYGrowth(1) & "     " & strOpCashFlowYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Capital Expenditures              " & dblAbsCapEx(0) & "     " & dblAbsCapEx(1) & "     " & dblAbsCapEx(2) & "     " & dblAbsCapEx(3) & Chr(10) & _
                "YOY Capital Expenditures Growth   " & strCapExYOYGrowth(0) & "     " & strCapExYOYGrowth(1) & "     " & strCapExYOYGrowth(2) & ""
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
'               YOY array -> YOY growth values
'                            YOY(0) is most recent year
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
            ScoreCashFlow = ScoreCashFlow + (CASH_FLOW_SCORE_MAX - i)
        End If
        YOYGrowth.Offset(0, i + 1) = YOY(i)
    Next i
    
    ScoreCashFlow = ScoreCashFlow * CASH_FLOW_SCORE_WEIGHT
    
    CheckCashFlowPassFail
    CashFlowScore
    
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

'===============================================================
' Procedure:    CashFlowScore
'
' Description:  Calculate score for cash flow
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  10Dec15 by Janice Laset Parkerson
'               - Initial Version
'===============================================================

Sub CashFlowScore()

    Range("FreeCashFlowScore") = ScoreCashFlow

End Sub
