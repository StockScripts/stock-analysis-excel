VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindStocksForm 
   Caption         =   "Find Stocks"
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
Option Explicit

'Stock Screener Parameters
Private Const MIN_PRICE = 5
Private Const AVG_VOLUME = 20000
Private Const CURRENT_RATIO_MIN = 1.5
Private Const LT_DEBT_TO_EQUITY_MAX = 35
Private Const ROE_MIN = 10


'===============================================================
' Procedure:    ButtonEnter_Click
'
' Description:  Uses Google Stock screener to find stocks
'               according to stock screener parameters. Prompts
'               user to enter inflation rate which determines the
'               minimum dividend yield requirement.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   09Sept14 by Janice Laset Parkerson
'                - Initial Version
'===============================================================
Public Sub ButtonEnter_Click()

    Dim strWebsite As String
    Dim dblDivYield As Double
    
    dblDivYield = FindStocksForm.TextBoxInflationRate.Text
    
    If FindStocksForm.TextBoxInflationRate.Text = "" Then
        MsgBox "You must enter a value."
        Exit Sub
    End If
    
    strWebsite = "http://www.google.com/finance/stockscreener#" & _
                    "c0=QuoteLast&min0=" & MIN_PRICE & "&" & _
                    "c1=AverageVolume&min1=" & AVG_VOLUME & "&" & _
                    "c2=DividendYield&min2=" & CURRENT_RATIO_MIN & "&" & _
                    "c3=CurrentRatioYear&min3=" & CurrentRatioMin & "&" & _
                    "c4=LTDebtToEquityYear&min4=0&max4=" & LT_DEBT_TO_EQUITY_MAX & "&" & _
                    "c5=ReturnOnEquityYear&min5=" & ROE_MIN & ""
         
    Unload FindStocksForm
    
    ActiveWorkbook.FollowHyperlink Address:=strWebsite

End Sub

'===============================================================
' Procedure     ButtonCancel_Click
'
' Description:  Closes Find Stocks form when user presses CANCEL
'               button.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   09Sept14 by Janice Laset Parkerson
'                - Initial Version
'===============================================================
Public Sub ButtonCancel_Click()

    Unload FindStocksForm
    
End Sub


