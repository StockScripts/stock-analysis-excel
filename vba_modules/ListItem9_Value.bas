Attribute VB_Name = "ListItem9_Value"
Option Explicit

Private ResultPrice As Result
Private Const PRICE_GROWTH_POTENTIAL_MIN = 0.2
Private Const PRICE_GROWTH_POTENTIAL_NEGATIVE = 0.1
'Public Const PRICE_SCORE_MAX = 20
Public ScorePrice As Integer

'===============================================================
' Procedure:    EvaluatePrice
'
' Description:  Display current price and target price information.
'               Calculate growth potential
'               if growth potential >= minimum required -> pass
'               else -> fail
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  22Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EvaluatePrice()

    Dim dblPriceGrowthPotential As Double
        
    ScorePrice = 0
    
    Range("Price").Offset(0, 1) = dblCurrentPrice
    Range("TargetPrice").Offset(0, 1) = dblTargetPrice
    Range("HighTarget").Offset(0, 1) = vHighTarget
    Range("LowTarget").Offset(0, 1) = vLowTarget
    Range("Brokers").Offset(0, 1) = vBrokers
    
    'calculate growth potential
    If dblTargetPrice = 0 Then
        dblPriceGrowthPotential = 0
    Else
        dblPriceGrowthPotential = (dblTargetPrice - dblCurrentPrice) / dblCurrentPrice
    End If
    
    Range("PriceGrowth").Offset(0, 1).Select
    If dblPriceGrowthPotential >= PRICE_GROWTH_POTENTIAL_MIN Then
        Selection.Font.ColorIndex = FONT_COLOR_GREEN
        ResultPrice = PASS
'        ScorePrice = PRICE_SCORE_MAX
    Else
        Selection.Font.ColorIndex = FONT_COLOR_RED
        ResultPrice = FAIL
'        If dblPriceGrowthPotential < PRICE_GROWTH_POTENTIAL_NEGATIVE Then
'            ScorePrice = ScorePrice - PRICE_SCORE_MAX
'        End If
    End If
    Range("PriceGrowth").Offset(0, 1) = dblPriceGrowthPotential
    
    DisplayPriceGrowthInfo
    
    CheckPricePassFail
'    PriceScore
    
End Sub

'===============================================================
' Procedure:    DisplayPriceGrowthInfo
'
' Description:  Comment box information for price growth
'               - price requirements
'               - target price information
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
' Rev History:  30Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub DisplayPriceGrowthInfo()

    Range("ListItemPrice") = "Is it overvalued?"
    
    'cell text
    Range("Price") = "Current Price"
    Range("TargetPrice") = "1 Yr Target Price"
    Range("PriceGrowth") = "Growth Potential"
    
    Range("HighTarget") = "High Target Price"
    Range("LowTarget") = "Low Target Price"
    Range("Brokers") = "Number of Brokers"
    
    'price growth info
    With Range("ListItemPrice")
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:="What is it:" & Chr(10) & _
                "   Target price is the projected price based on earnings forecast and valuation" & Chr(10) & _
                "   as stated by investment analysts." & Chr(10) & _
                "Why is it important:" & Chr(10) & _
                "   The target price aids in evaluating the potential risk or reward for a stock." & Chr(10) & _
                "   This should be taken into account because analyst estimates influence the stock price." & Chr(10) & _
                "What to look for:" & Chr(10) & _
                "   To account for error, the target price should be at least 20% higher than the current price." & Chr(10) & _
                "What to watch for:" & Chr(10) & _
                "   If the company offers dividends, this percentage can be added to the potential growth of a stock."
        .Comment.Shape.TextFrame.AutoSize = True
    End With

End Sub

'===============================================================
' Procedure:    CheckPricePassFail
'
' Description:  Display check or x mark if the price
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
' Rev History:  22Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub CheckPricePassFail()

    If ResultPrice = PASS Then
        Range("PriceCheck") = CHECK_MARK
        Range("PriceCheck").Font.ColorIndex = FONT_COLOR_GREEN
    Else
        Range("PriceCheck") = X_MARK
        Range("PriceCheck").Font.ColorIndex = FONT_COLOR_RED
    End If
    
End Sub

'===============================================================
' Procedure:    PriceScore
'
' Description:  Calculate score for price
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

Sub PriceScore()

'    Range("PriceScore") = ScorePrice

End Sub
