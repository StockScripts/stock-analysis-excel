Attribute VB_Name = "Item7_RedFlags"
Private Const RedFlagMaxIncrease = 0.2

Public ReceivablesToSales1 As Double
Public ReceivablesToSales2 As Double
Public ReceivablesToSales3 As Double
Public ReceivablesToSales4 As Double
Public ReceivablesToSales5 As Double

Public InventoryToSales1 As Double
Public InventoryToSales2 As Double
Public InventoryToSales3 As Double
Public InventoryToSales4 As Double
Public InventoryToSales5 As Double

Public SGAToSales1 As Double
Public SGAToSales2 As Double
Public SGAToSales3 As Double
Public SGAToSales4 As Double
Public SGAToSales5 As Double

Sub RedFlags()

    ReceivablesToSales
    InventoryToSales
    SGAToSales

End Sub

Sub ReceivablesToSales()

    Dim ErrorNum As Integer
    
    On Error GoTo ErrorHandler

    Range("A26").Font.Bold = True
    Range("A26") = "Are there any red flags?"
    
'   name Revenue cell
    Range("B27").Name = "ReceivablesToSales"
    Range("27:27").Name = "ReceivablesToSalesRow"
    
    Range("ReceivablesToSalesRow").NumberFormat = "0.0%"
    
'   write "Revenue" text
    Range("ReceivablesToSales").HorizontalAlignment = xlLeft
    Range("ReceivablesToSales") = "Receivables/Sales"
        
'   populate ROE information
    ErrorNum = 1
    ReceivablesToSales1 = Receivables1 / Revenue1
    Range("ReceivablesToSales").Offset(0, 1) = ReceivablesToSales1
    
    ErrorNum = 2
    ReceivablesToSales2 = Receivables2 / Revenue2
    Range("ReceivablesToSales").Offset(0, 2) = ReceivablesToSales2
    
    ErrorNum = 3
    ReceivablesToSales3 = Receivables3 / Revenue3
    Range("ReceivablesToSales").Offset(0, 3) = ReceivablesToSales3
    
    ErrorNum = 4
    ReceivablesToSales4 = Receivables4 / Revenue4
    Range("ReceivablesToSales").Offset(0, 4) = ReceivablesToSales4
    
    ErrorNum = 5
    ReceivablesToSales5 = Receivables5 / Revenue5
    Range("ReceivablesToSales").Offset(0, 5) = ReceivablesToSales5

'   calculate YOY Debt To Equity growth
    ReceivablesToSalesYOY
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case 1
            ReceivablesToSales1 = 0
            Range("ReceivablesToSales").Offset(0, 1) = ReceivablesToSales1
        Case 2
            ReceivablesToSales2 = 0
            Range("ReceivablesToSales").Offset(0, 2) = ReceivablesToSales2
        Case 3
            ReceivablesToSales3 = 0
            Range("ReceivablesToSales").Offset(0, 3) = ReceivablesToSales3
        Case 4
            ReceivablesToSales4 = 0
            Range("ReceivablesToSales").Offset(0, 4) = ReceivablesToSales4
        Case 5
            ReceivablesToSales5 = 0
            Range("ReceivablesToSales").Offset(0, 5) = ReceivablesToSales5
   End Select
   
   ReceivablesToSalesYOY

End Sub

Sub ReceivablesToSalesYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B28").Name = "YOYGrowth"
    Range("28:28").Name = "YOYRow"
    
    Range("YOYRow").NumberFormat = "0.0%"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(ReceivablesToSales1, ReceivablesToSales2)
    YOY2 = YOYGrowth(ReceivablesToSales2, ReceivablesToSales3)
    YOY3 = YOYGrowth(ReceivablesToSales3, ReceivablesToSales4)
    YOY4 = YOYGrowth(ReceivablesToSales4, ReceivablesToSales5)
    
    Call YOYRedFlagEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub


Sub InventoryToSales()

    Dim ErrorNum As Integer
    
    On Error GoTo ErrorHandler

'   name Revenue cell
    Range("B29").Name = "InventoryToSales"
    Range("29:29").Name = "InventoryToSalesRow"
    
    Range("InventoryToSalesRow").NumberFormat = "0.0%"
    
'   write "Revenue" text
    Range("InventoryToSales").HorizontalAlignment = xlLeft
    Range("InventoryToSales") = "Inventory/Sales"
    
'   populate ROE information
    ErrorNum = 1
    InventoryToSales1 = Inventory1 / Revenue1
    Range("InventoryToSales").Offset(0, 1) = InventoryToSales1
    
    ErrorNum = 2
    InventoryToSales2 = Inventory2 / Revenue2
    Range("InventoryToSales").Offset(0, 2) = InventoryToSales2
    
    ErrorNum = 3
    InventoryToSales3 = Inventory3 / Revenue3
    Range("InventoryToSales").Offset(0, 3) = InventoryToSales3
    
    ErrorNum = 4
    InventoryToSales4 = Inventory4 / Revenue4
    Range("InventoryToSales").Offset(0, 4) = InventoryToSales4
    
    ErrorNum = 5
    InventoryToSales5 = Inventory5 / Revenue5
    Range("InventoryToSales").Offset(0, 5) = InventoryToSales5

'   calculate YOY Debt To Equity growth
    InventoryToSalesYOY
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case 1
            InventoryToSales1 = 0
            Range("InventoryToSales").Offset(0, 1) = InventoryToSales1
        Case 2
            InventoryToSales2 = 0
            Range("InventoryToSales").Offset(0, 2) = InventoryToSales2
        Case 3
            InventoryToSales3 = 0
            Range("InventoryToSales").Offset(0, 3) = InventoryToSales3
        Case 4
            InventoryToSales4 = 0
            Range("InventoryToSales").Offset(0, 4) = InventoryToSales4
        Case 5
            InventoryToSales5 = 0
            Range("InventoryToSales").Offset(0, 5) = InventoryToSales5
   End Select
   
   InventoryToSalesYOY

End Sub

Sub InventoryToSalesYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B30").Name = "YOYGrowth"
    Range("30:30").Name = "YOYRow"
    
    Range("YOYRow").NumberFormat = "0.0%"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(InventoryToSales1, InventoryToSales2)
    YOY2 = YOYGrowth(InventoryToSales2, InventoryToSales3)
    YOY3 = YOYGrowth(InventoryToSales3, InventoryToSales4)
    YOY4 = YOYGrowth(InventoryToSales4, InventoryToSales5)
    
    Call YOYRedFlagEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub


Sub SGAToSales()

    Dim ErrorNum As Integer
    
    On Error GoTo ErrorHandler

'   name Revenue cell
    Range("B31").Name = "SGAToSales"
    Range("31:31").Name = "SGAToSalesRow"
    
    Range("SGAToSalesRow").NumberFormat = "0.0%"
    
'   write "Revenue" text
    Range("SGAToSales").HorizontalAlignment = xlLeft
    Range("SGAToSales") = "SGA/Sales"
    
'   populate ROE information
    ErrorNum = 1
    SGAToSales1 = SGA1 / Revenue1
    Range("SGAToSales").Offset(0, 1) = SGAToSales1
    
    ErrorNum = 2
    SGAToSales2 = SGA2 / Revenue2
    Range("SGAToSales").Offset(0, 2) = SGAToSales2
    
    ErrorNum = 3
    SGAToSales3 = SGA3 / Revenue3
    Range("SGAToSales").Offset(0, 3) = SGAToSales3
    
    ErrorNum = 4
    SGAToSales4 = SGA4 / Revenue4
    Range("SGAToSales").Offset(0, 4) = SGAToSales4
    
    ErrorNum = 5
    SGAToSales5 = SGA5 / Revenue5
    Range("SGAToSales").Offset(0, 5) = SGAToSales5

    Range("SGAToSales").AddComment
    Range("SGAToSales").Comment.Visible = False
    Range("SGAToSales").Comment.Text Text:="Overhead costs" & Chr(10) & _
                "operating expenses except cost of sales, " & Chr(10) & _
                "R&D, and depreciation and amortization." & Chr(10) & _
                "can be used to detect operational problems along with deteriorating operating margins" & Chr(10) & _
                "SGA/Sales should be stable and not increasing"
    Range("SGAToSales").Comment.Shape.TextFrame.AutoSize = True
    
    SGAToSalesYOY
    Exit Sub
    
ErrorHandler:

    Select Case ErrorNum
        Case 1
            SGAToSales1 = 0
            Range("SGAToSales").Offset(0, 1) = SGAToSales1
        Case 2
            SGAToSales2 = 0
            Range("SGAToSales").Offset(0, 2) = SGAToSales2
        Case 3
            SGAToSales3 = 0
            Range("SGAToSales").Offset(0, 3) = SGAToSales3
        Case 4
            SGAToSales4 = 0
            Range("SGAToSales").Offset(0, 4) = SGAToSales4
        Case 5
            SGAToSales5 = 0
            Range("SGAToSales").Offset(0, 5) = SGAToSales5
   End Select
   
   SGAToSalesYOY

End Sub

Sub SGAToSalesYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B32").Name = "YOYGrowth"
    Range("32:32").Name = "YOYRow"
    
    Range("YOYRow").NumberFormat = "0.0%"
     
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(SGAToSales1, SGAToSales2)
    YOY2 = YOYGrowth(SGAToSales2, SGAToSales3)
    YOY3 = YOYGrowth(SGAToSales3, SGAToSales4)
    YOY4 = YOYGrowth(SGAToSales4, SGAToSales5)
    
    Call YOYRedFlagEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYRedFlagEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    If YOY4 > RedFlagMaxIncrease Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    If YOY3 > RedFlagMaxIncrease Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
        If YOY2 > RedFlagMaxIncrease Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If YOY1 > RedFlagMaxIncrease Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function




