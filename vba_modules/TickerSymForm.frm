VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TickerSymForm 
   Caption         =   "Ticker Symbol"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4845
   OleObjectBlob   =   "TickerSymForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TickerSymForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================
' Procedure:    ButtonCreateReport_Click
'
' Description:  Calls AnalyzeStock procedure to begin stock
'               analysis of the company entered.
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
Private Sub ButtonCreateReport_Click()

    If TickerSymForm.TextBoxTickerSym.Text = "" Then
        MsgBox "You must enter a Ticker Symbol."
        Exit Sub
    End If
        
    TickerSym = TickerSymForm.TextBoxTickerSym.Text
         
    Unload TickerSymForm
    AnalyzeStock
    
End Sub

'===============================================================
' Procedure:    ButtonCancel_Click
'
' Description:  Closes Stock Analysis form when user presses
'               CANCEL button.
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
Private Sub ButtonCancel_Click()

    Unload TickerSymForm
    
End Sub
