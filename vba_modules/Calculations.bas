Attribute VB_Name = "Calculations"
Option Explicit

'calculate year over year growth for financials
Function YOYGrowth(RecentYear, PastYear)
    On Error Resume Next
    If PastYear = 0 Then
        YOYGrowth = 0
    Else
        YOYGrowth = (RecentYear - PastYear) / Abs(PastYear)
    End If
    
End Function

Function YOYEvaluation(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    Selection.NumberFormat = "0.0%"
    If YOY4 < 0 Then
        Selection.Font.ColorIndex = 3   'red font
    Else
        Selection.Font.ColorIndex = 10   'green font
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    Selection.NumberFormat = "0.0%"
    If YOY3 < 0 Then
        Selection.Font.ColorIndex = 3   'red font
    ElseIf YOY3 > YOY4 Then
        Selection.Font.ColorIndex = 10   'green font
    Else
        Selection.Font.ColorIndex = 46   'orange font
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    Selection.NumberFormat = "0.0%"
    If YOY2 < 0 Then
        Selection.Font.ColorIndex = 3   'red font
    ElseIf YOY2 > YOY3 Then
        Selection.Font.ColorIndex = 10   'green font
    Else
        Selection.Font.ColorIndex = 46   'orange font
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    Selection.NumberFormat = "0.0%"
    If YOY1 < 0 Then
        Selection.Font.ColorIndex = 3   'red font
    ElseIf YOY1 > YOY2 Then
        Selection.Font.ColorIndex = 10   'green font
    Else
        Selection.Font.ColorIndex = 46   'orange font
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function


Function YOYDecrease(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    Selection.NumberFormat = "0.0%"
    If YOY4 < 0 Then
        Selection.Font.ColorIndex = 10   'green font
    Else
        Selection.Font.ColorIndex = 3   'red font
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    Selection.NumberFormat = "0.0%"
    If YOY3 < 0 Then
        Selection.Font.ColorIndex = 10   'green font
    ElseIf YOY3 > YOY4 Then
        Selection.Font.ColorIndex = 3   'red font
    Else
        Selection.Font.ColorIndex = 46   'orange font
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    Selection.NumberFormat = "0.0%"
    If YOY2 < 0 Then
        Selection.Font.ColorIndex = 10   'green font
    ElseIf YOY2 > YOY3 Then
        Selection.Font.ColorIndex = 3   'red font
    Else
        Selection.Font.ColorIndex = 46   'orange font
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    Selection.NumberFormat = "0.0%"
    If YOY1 < 0 Then
        Selection.Font.ColorIndex = 10   'green font
    ElseIf YOY1 > YOY2 Then
        Selection.Font.ColorIndex = 3   'red font
    Else
        Selection.Font.ColorIndex = 46   'orange font
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function
