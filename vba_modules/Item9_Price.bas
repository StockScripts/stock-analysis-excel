Attribute VB_Name = "Item9_Price"

Sub PriceToEarnings()

'   name Revenue cell
    Range("B37").Name = "PricePerShare"
    
'   name EPS cell
    Range("B24").Name = "DilutedEPS"
    
'   write "Price/Share" text
    Range("PricePerShare").HorizontalAlignment = xlLeft
    Range("PricePerShare") = "Enter Price/Share"
    
'   name P/E cell
    Range("B38").Name = "PricePerEarnings"
    
'   write "Price/Share" text
    Range("PricePerEarnings").HorizontalAlignment = xlLeft
    Range("PricePerEarnings") = "Price/Earnings"

    Range("PricePerEarnings").Offset(0, 1).FormulaR1C1 = "=R[-1]C/R[-14]C"
    Range("PricePerEarnings").Offset(0, 2).FormulaR1C1 = "=R[-1]C/R[-14]C"
    Range("PricePerEarnings").Offset(0, 3).FormulaR1C1 = "=R[-1]C/R[-14]C"
    Range("PricePerEarnings").Offset(0, 4).FormulaR1C1 = "=R[-1]C/R[-14]C"
    Range("PricePerEarnings").Offset(0, 5).FormulaR1C1 = "=R[-1]C/R[-14]C"
    
End Sub
