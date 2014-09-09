VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindStocksForm 
   Caption         =   "Dividend Yield"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FindStocksForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FindStocksForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const MinPrice = 5
Private Const AvgVolume = 20000
Private Const CurrentRatioMin = 1.5
Private Const LTDebtToEquityMax = 35
Private Const ROEMin = 10

Public Sub CancelButton_Click()
    Unload FindStocksForm
End Sub

Public Sub EnterButton_Click()

    Dim strWebsite As String
    Dim DivYield As Double
    
    DivYield = FindStocksForm.TextBoxInflationRate.Text
    
    If FindStocksForm.TextBoxInflationRate.Text = "" Then
        MsgBox "You must enter a value."
        Exit Sub
    End If
    
'    strWebsite = "http://www.google.com/finance/stockscreener#c0=QuoteLast&min0=" & MinPrice & "&max0=" & MaxPrice & "&c1=ReturnOnInvestment5Years&min1=10&c2=ReturnOnAssets5Years&min2=10&c3=ReturnOnEquity5Years&min3=10&c4=NetIncomeGrowthRate5Years&min4=10&c5=RevenueGrowthRate5Years&min5=10&region=us&sector=AllSectors&sort=&sortOrder="
    strWebsite = "http://www.google.com/finance/stockscreener#" & _
                    "c0=QuoteLast&min0=" & MinPrice & "&" & _
                    "c1=AverageVolume&min1=" & AvgVolume & "&" & _
                    "c2=DividendYield&min2=" & DivYield & "&" & _
                    "c3=CurrentRatioYear&min3=" & CurrentRatioMin & "&" & _
                    "c4=LTDebtToEquityYear&min4=0&max4=" & LTDebtToEquityMax & "&" & _
                    "c5=ReturnOnEquityYear&min5=" & ROEMin & ""
         
    Unload FindStocksForm
    
    ActiveWorkbook.FollowHyperlink Address:=strWebsite

End Sub

