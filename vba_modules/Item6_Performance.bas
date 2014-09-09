Attribute VB_Name = "Item6_Performance"

Sub EPS()

    Range("A23").Font.Bold = True
    Range("A23") = "Have they been performing well?"
    
'   name EPS cell
    Range("B24").Name = "DilutedEPS"
    
'   write "Diluted EPS" text
    Range("DilutedEPS").HorizontalAlignment = xlLeft
    Range("DilutedEPS") = "Diluted EPS"
    
'   populate revenue information
    If EPS1 >= 0 Then
        Range("DilutedEPS").Offset(0, 1).Font.ColorIndex = GreenFont
    Else
        Range("DilutedEPS").Offset(0, 1).Font.ColorIndex = RedFont
    End If
    Range("DilutedEPS").Offset(0, 1) = EPS1
    
    If EPS2 >= 0 Then
        Range("DilutedEPS").Offset(0, 2).Font.ColorIndex = GreenFont
    Else
        Range("DilutedEPS").Offset(0, 2).Font.ColorIndex = RedFont
    End If
    Range("DilutedEPS").Offset(0, 2) = EPS2
    
    If EPS3 >= 0 Then
        Range("DilutedEPS").Offset(0, 3).Font.ColorIndex = GreenFont
    Else
        Range("DilutedEPS").Offset(0, 3).Font.ColorIndex = RedFont
    End If
    Range("DilutedEPS").Offset(0, 3) = EPS3
    
    If EPS4 >= 0 Then
        Range("DilutedEPS").Offset(0, 4).Font.ColorIndex = GreenFont
    Else
        Range("DilutedEPS").Offset(0, 4).Font.ColorIndex = RedFont
    End If
    Range("DilutedEPS").Offset(0, 4) = EPS4
    
    If EPS5 >= 0 Then
        Range("DilutedEPS").Offset(0, 5).Font.ColorIndex = GreenFont
    Else
        Range("DilutedEPS").Offset(0, 5).Font.ColorIndex = RedFont
    End If
    Range("DilutedEPS").Offset(0, 5) = EPS5
    
    Range("DilutedEPS").AddComment
    Range("DilutedEPS").Comment.Visible = False
    Range("DilutedEPS").Comment.Text Text:="earnings = sales x profit margin" & Chr(10) & _
                "to increase earnings"
    Range("DilutedEPS").Comment.Shape.TextFrame.AutoSize = True
    
'   calculate YOY revenue growth
    EPSYOY

End Sub

Sub EPSYOY()

    Dim YOY1 As Double
    Dim YOY2 As Double
    Dim YOY3 As Double
    Dim YOY4 As Double
    Dim YOY5 As Double
    
'   name YOY cell
    Range("B25").Name = "YOYGrowth"
    Range("25:25").Name = "YOYRow"
    
'   write "YOY Growth" text
    Range("YOYGrowth").HorizontalAlignment = xlRight
    Range("YOYGrowth") = "YOY Growth (%)"
    
    Range("YOYRow").NumberFormat = "0.0%"
    
    Range("YOYRow").Font.Italic = True
    With Range("YOYRow").Font
        .Color = -6908266
        .TintAndShade = 0
    End With

'   populate YOY growth information
    
    YOY1 = YOYGrowth(EPS1, EPS2)
    YOY2 = YOYGrowth(EPS2, EPS3)
    YOY3 = YOYGrowth(EPS3, EPS4)
    YOY4 = YOYGrowth(EPS4, EPS5)
    
    Call YOYEPSEval(Range("YOYGrowth"), YOY1, YOY2, YOY3, YOY4)
    
    Range("YOYGrowth").Offset(0, 5).HorizontalAlignment = xlCenter
    Range("YOYGrowth").Offset(0, 5) = "---"
    
End Sub

Function YOYEPSEval(YOYGrowth As Range, YOY1, YOY2, YOY3, YOY4)
    
    YOYGrowth.Offset(0, 4).Select
    If EPS4 < 0 Or YOY4 < 0 Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 4) = YOY4
    
    YOYGrowth.Offset(0, 3).Select
    If EPS3 < 0 Or YOY3 < 0 Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 3) = YOY3
    
    YOYGrowth.Offset(0, 2).Select
    If EPS2 < 0 Or YOY2 < 0 Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 2) = YOY2
    
    YOYGrowth.Offset(0, 1).Select
    If EPS1 < 0 Or YOY1 < 0 Then
        Selection.Font.ColorIndex = RedFont
    Else
        Selection.Font.ColorIndex = GreenFont
    End If
    YOYGrowth.Offset(0, 1) = YOY1
    
End Function


