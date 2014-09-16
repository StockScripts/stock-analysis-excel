VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormFindStocks 
   Caption         =   "Find Stocks"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FormFindStocks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormFindStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Stock Screener Parameters
Private Const MIN_PRICE = 2
Private Const AVG_VOLUME = 20000
Private Const CURRENT_RATIO_MIN = 1.5
Private Const LT_DEBT_TO_EQUITY_MAX = 40
Private Const ROE_MIN = 10

'Stock Screener Size Parameters
Private Const SMALL_CAP_MIN = 250000000
Private Const SMALL_CAP_MAX = 2000000000
Private Const MID_CAP_MIN = 2000000000
Private Const MID_CAP_MAX = 10000000000#
Private Const LARGE_CAP_MIN = 10000000000#

Private OptionSmallCap As Boolean
Private OptionMidCap As Boolean
Private OptionLargeCap As Boolean

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
    
    If OptionSmallCap Then
        strWebsite = "http://www.google.com/finance/stockscreener#" & _
                    "c0=MarketCap&min0=" & SMALL_CAP_MIN & "&max0=" & SMALL_CAP_MAX & "&" & _
                    "c1=QuoteLast&min1=" & MIN_PRICE & "&" & _
                    "c2=AverageVolume&min2=" & AVG_VOLUME & "&" & _
                    "c3=CurrentRatioYear&min3=" & CURRENT_RATIO_MIN & "&" & _
                    "c4=LTDebtToEquityYear&min4=0&max4=" & LT_DEBT_TO_EQUITY_MAX & "&" & _
                    "c5=ReturnOnEquityYear&min5=" & ROE_MIN & ""
         
    ElseIf OptionMidCap Then
        strWebsite = "http://www.google.com/finance/stockscreener#" & _
                    "c0=MarketCap&min0=" & MID_CAP_MIN & "&max0=" & MID_CAP_MAX & "&" & _
                    "c1=QuoteLast&min1=" & MIN_PRICE & "&" & _
                    "c2=AverageVolume&min2=" & AVG_VOLUME & "&" & _
                    "c3=CurrentRatioYear&min3=" & CURRENT_RATIO_MIN & "&" & _
                    "c4=LTDebtToEquityYear&min4=0&max4=" & LT_DEBT_TO_EQUITY_MAX & "&" & _
                    "c5=ReturnOnEquityYear&min5=" & ROE_MIN & ""
    Else
        strWebsite = "http://www.google.com/finance/stockscreener#" & _
                    "c0=MarketCap&min0=" & LARGE_CAP_MIN & "&" & _
                    "c1=QuoteLast&min1=" & MIN_PRICE & "&" & _
                    "c2=AverageVolume&min2=" & AVG_VOLUME & "&" & _
                    "c3=CurrentRatioYear&min3=" & CURRENT_RATIO_MIN & "&" & _
                    "c4=LTDebtToEquityYear&min4=0&max4=" & LT_DEBT_TO_EQUITY_MAX & "&" & _
                    "c5=ReturnOnEquityYear&min5=" & ROE_MIN & ""
    End If
                    
    Unload FormFindStocks
    
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

'===============================================================
' Procedure     OptionButtonSmallCap_Click
'
' Description:  Sets flag to use small cap range using Google
'               Stock Screener
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   15Sept14 by Janice Laset Parkerson
'                - Initial Version
'===============================================================
Public Sub OptionButtonSmallCap_Click()

    OptionSmallCap = True
    OptionMidCap = False
    OptionLargeCap = False
    
End Sub

'===============================================================
' Procedure     OptionButtonMidCap_Click
'
' Description:  Sets flag to use mid cap range using Google
'               Stock Screener
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   15Sept14 by Janice Laset Parkerson
'                - Initial Version
'===============================================================
Public Sub OptionButtonMidCap_Click()

    OptionSmallCap = False
    OptionMidCap = True
    OptionLargeCap = False
    
End Sub

'===============================================================
' Procedure     OptionButtonLargeCap_Click
'
' Description:  Sets flag to use large cap range using Google
'               Stock Screener
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   15Sept14 by Janice Laset Parkerson
'                - Initial Version
'===============================================================
Public Sub OptionButtonLargeCap_Click()

    OptionSmallCap = False
    OptionMidCap = False
    OptionLargeCap = True
    
End Sub
