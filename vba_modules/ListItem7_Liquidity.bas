Attribute VB_Name = "ListItem7_Liquidity"
Option Explicit

Private dblQuickRatio(0 To 4) As Double

Private Const QUICK_RATIO_MIN = 1
Private ResultLiquidity As Result

'===============================================================
' Procedure:    EvaluateQuickRatio
'
' Description:  Call procedure to display liquidity information.
'               Call procedure to display YOY growth information
'               if recent year quick ratio > required min -> pass
'               else -> fail
'
'               if past year quick ratio > required min -> warning
'
'               catch errors and set value to STR_NO_DATA
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  19ept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluateQuickRatio()

    Dim i As Integer
    
    On Error Resume Next

    ResultLiquidity = PASS
    
    DisplayLiquidityInfo
    
    dblQuickRatio(0) = (dblCurrentAssets(0) - dblInventory(0)) / dblCurrentLiabilities(0)
    Range("QuickRatio").Offset(0, 1).Select
    If Err Then
        Selection.HorizontalAlignment = xlCenter
        Selection.Value = STR_NO_DATA
        Err.Clear
    Else
        If dblQuickRatio(0) >= QUICK_RATIO_MIN Then
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
        Else
            Selection.Font.ColorIndex = FONT_COLOR_RED
            ResultLiquidity = FAIL
        End If
        Selection.Value = dblQuickRatio(0)
    End If
    
    For i = 1 To (iYearsAvailableIncome - 1)
        dblQuickRatio(i) = (dblCurrentAssets(i) - dblInventory(i)) / dblCurrentLiabilities(0)
        Range("QuickRatio").Offset(0, i + 1).Select
        If Err Then
            Selection.HorizontalAlignment = xlCenter
            Selection.Value = STR_NO_DATA
            Err.Clear
        Else
            If dblQuickRatio(i) >= QUICK_RATIO_MIN Then
                Selection.Font.ColorIndex = FONT_COLOR_GREEN
            Else
                Selection.Font.ColorIndex = FONT_COLOR_ORANGE
                ResultLiquidity = FAIL
            End If
            Selection.Value = dblQuickRatio(i)
        End If
    Next i

    CalculateQuickRatioYOYGrowth

End Sub

'===============================================================
' Procedure:    DisplayLiquidityInfo
'
' Description:  Comment box information for Liquidity
'               - quick ratio requirements
'               - quick ratio info
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  11Nov14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub DisplayLiquidityInfo()

    Dim dblCurrentAssetsYOYGrowth(0 To 3) As Double
    Dim strCurrentAssetsYOYGrowth(0 To 3) As String
    
    Dim dblInventoryYOYGrowth(0 To 3) As Double
    Dim strInventoryYOYGrowth(0 To 3) As String
    
    Dim dblCurrentLiabilitiesYOYGrowth(0 To 3) As Double
    Dim strCurrentLiabilitiesYOYGrowth(0 To 3) As String
    
    Dim i As Integer
    
    Range("ListItemQuickRatio") = "Are debts covered?"
    Range("QuickRatio") = "Quick Ratio"
    
    With Range("ListItemQuickRatio")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="What is it:" & Chr(10) & _
                "   Quick ratio is used to gauge a company's liquidity. It is a better measure than " & Chr(10) & _
                "   current ratio because it excludes inventory, which takes time to be converted to cash." & Chr(10) & _
                "Why is it important:" & Chr(10) & _
                "   Quick ratio measures the company's ability to pay their short term obligations." & Chr(10) & _
                "What to look for:" & Chr(10) & _
                "   Quick ratio should be greater than 1 and not decreasing." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   A high ratio means a better position for the company, but it also means the cash is not being used for growth." & Chr(10) & _
                "   A company with a quick ratio of less than 1 cannot currently pay its current liabilities."
        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    'calculate YOY growth
    For i = 0 To (iYearsAvailableIncome - 2)
        dblCurrentAssetsYOYGrowth(i) = CalculateYOYGrowth(dblCurrentAssets(i), dblCurrentAssets(i + 1))
        strCurrentAssetsYOYGrowth(i) = Format(dblCurrentAssetsYOYGrowth(i), "0.0%")
        
        dblInventoryYOYGrowth(i) = CalculateYOYGrowth(dblInventory(i), dblInventory(i + 1))
        strInventoryYOYGrowth(i) = Format(dblInventoryYOYGrowth(i), "0.0%")
        
        dblCurrentLiabilitiesYOYGrowth(i) = CalculateYOYGrowth(dblCurrentLiabilities(i), dblCurrentLiabilities(i + 1))
        strCurrentLiabilitiesYOYGrowth(i) = Format(dblCurrentLiabilitiesYOYGrowth(i), "0.0%")
    Next i
    
    With Range("QuickRatio")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="Quick Ratio = (Current Assets - Inventory) / Current Liabilities" & Chr(10) & _
                "" & Chr(10) & _
                "YOY Current Assets" & "                " & dblCurrentAssets(0) & "      " & dblCurrentAssets(1) & "      " & dblCurrentAssets(2) & "      " & dblCurrentAssets(3) & Chr(10) & _
                "YOY Current Assets Growth     " & strCurrentAssetsYOYGrowth(0) & "     " & strCurrentAssetsYOYGrowth(1) & "     " & strCurrentAssetsYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Inventory" & "                " & dblInventory(0) & "      " & dblInventory(1) & "      " & dblInventory(2) & "      " & dblInventory(3) & Chr(10) & _
                "YOY Inventory Growth     " & strInventoryYOYGrowth(0) & "     " & strInventoryYOYGrowth(1) & "     " & strInventoryYOYGrowth(2) & Chr(10) & _
                "" & Chr(10) & _
                "YOY Current Liabilities              " & dblCurrentLiabilities(0) & "      " & dblCurrentLiabilities(1) & "      " & dblCurrentLiabilities(2) & "      " & dblCurrentLiabilities(3) & Chr(10) & _
                "YOY Current Liabilities Growth   " & strCurrentLiabilitiesYOYGrowth(0) & "     " & strCurrentLiabilitiesYOYGrowth(1) & "     " & strCurrentLiabilitiesYOYGrowth(2) & ""
        .Comment.Shape.TextFrame.AutoSize = True
    End With

End Sub

'===============================================================
' Procedure:    CalculateQuickRatioYOYGrowth
'
' Description:  Call procedure to calculate and display YOY
'               growth for quick ratio data. Format cells.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  19Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CalculateQuickRatioYOYGrowth()

    Dim dblYOYGrowth(0 To 3) As Double
    Dim i As Integer
    
    Range("QuickRatioYOYGrowth") = "YOY Growth (%)"

    'populate YOY growth information
    '(0) is most recent year
    For i = 0 To (iYearsAvailableIncome - 2)
        dblYOYGrowth(i) = CalculateYOYGrowth(dblQuickRatio(i), dblQuickRatio(i + 1))
        
        If Err Then
            dblYOYGrowth(i) = 0
            Err.Clear
        End If
    Next i
    
    Call EvaluateQuickRatioYOYGrowth(Range("QuickRatioYOYGrowth"), dblYOYGrowth)
    
End Sub

'===============================================================
' Procedure:    EvaluateQuickRatioYOYGrowth
'
' Description:  Display YOY growth information.
'               for the most recent year
'                   if quick ratio is less than the min -> red font
'                   else if quick ratio is decreasing -> orange font
'                   else quick ratio is increasing -> green font
'               for previous years
'                   if quick ratio is less than the min or decreasing -> orange font
'                   else quick ratio is increasing -> green font
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   YOYGrowth As Range -> first cell of net margin YOY growth
'               YOY array -> YOY growth values
'                            YOY(0) is most recent year
'
' Returns:      N/A
'
' Rev History:  19Sept14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Function EvaluateQuickRatioYOYGrowth(YOYGrowth As Range, YOY() As Double)
    
    Dim i As Integer
    
    YOYGrowth.Offset(0, 1).Select
    If dblQuickRatio(0) < QUICK_RATIO_MIN Then
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultLiquidity = FAIL
    ElseIf YOY(0) < 0 Then
        Selection.Font.ColorIndex = FONT_COLOR_ORANGE
    Else
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
    End If
    Selection.Value = YOY(0)
    
    For i = 1 To (iYearsAvailableIncome - 2)
        YOYGrowth.Offset(0, i + 1).Select
        If dblQuickRatio(i) < QUICK_RATIO_MIN Or YOY(i) < 0 Then
            Selection.Font.ColorIndex = FONT_COLOR_ORANGE
        Else
            Selection.Font.ColorIndex = FONT_COLOR_GREEN
        End If
        YOYGrowth.Offset(0, i + 1) = YOY(i)
    Next i
    
    CheckLiquidityPassFail
    
End Function

'===============================================================
' Procedure:    CheckLiquidityPassFail
'
' Description:  Display check or x mark if the leverage
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
' Rev History:  14Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CheckLiquidityPassFail()

    If ResultLiquidity = PASS Then
        Range("LiquidityCheck") = CHECK_MARK
        Range("LiquidityCheck").Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("LiquidityCheck") = X_MARK
        Range("LiquidityCheck").Font.ColorIndex = FONT_COLOR_RED
    End If

End Sub


