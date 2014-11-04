Attribute VB_Name = "DataImport"
Option Explicit

Global ie As InternetExplorer
Global htmlFinancialStatement As HTMLDocument

Enum READYSTATE
    READYSTATE_UNINITIALIZED = 0
    READYSTATE_LOADING = 1
    READYSTATE_LOADED = 2
    READYSTATE_INTERACTIVE = 3
    READYSTATE_COMPLETE = 4
End Enum

'===============================================================
' Procedure:    ImportFinancialData
'
' Description:  Call procedures to open Google Finance page, enter
'               ticker symbol, and get financial data.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   07Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub ImportFinancialData()
   
   OpenGoogleFinance
   EnterTickerSymbol
   GetFinancialStatements
    
End Sub

'===============================================================
' Procedure:    OpenGoogleFinance
'
' Description:  Launch Internet Explorer and navigate to Google
'               Finance page.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   07Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub OpenGoogleFinance()

    Set ie = New InternetExplorer
    ie.Visible = False
    ie.navigate "google.com/finance"

    'Wait until IE is done loading page
    Do While ie.READYSTATE <> READYSTATE_COMPLETE
        Application.StatusBar = "Navigating to Google Finance ..."
        DoEvents
    Loop
    
    'get text of HTML document returned
    Set htmlFinancialStatement = ie.document
    
    'reset status bar
    Application.StatusBar = ""
End Sub

'===============================================================
' Procedure:    EnterTickerSymbol
'
' Description:  Find ticker symbol input field and enter ticker
'               symbol. Find search button and click button.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   07Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub EnterTickerSymbol()

    Dim ObjCollection As IHTMLElementCollection
    Dim Object As IHTMLElement
    
    Dim r As Long
        
    'get input tags on page
    Set ObjCollection = htmlFinancialStatement.getElementsByTagName("input")
    
    'ticker symbol entry is child(2) of input element collection
    ObjCollection(2).Value = strTickerSym
    
    'get ticker symbol search button
    Set ObjCollection = htmlFinancialStatement.getElementsByTagName("button")
    
    'search button is child(0) of button element collection
    ObjCollection(0).Click
    
    'Wait until IE is done loading new page
    Do While ie.READYSTATE <> READYSTATE_INTERACTIVE
        Application.StatusBar = "Navigating to company page..."
        DoEvents
    Loop
    
    Do While ie.READYSTATE <> READYSTATE_COMPLETE
        Application.StatusBar = "Navigating to company page..."
        DoEvents
    Loop

End Sub

'===============================================================
' Procedure:    GetFinancialStatements
'
' Description:  Ensure company page is loaded correctly.  Find
'               link for financial statments and click link.
'               Ensure financials page is loaded correctly.
'
' Author:       Janice Laset Parkerson
'
' Notes:        N/A
'
' Parameters:   N/A
'
' Returns:      N/A
'
'Rev History:   07Oct14 by Janice Laset Parkerson
'               - Initial Version
'===============================================================
Sub GetFinancialStatements()

    Dim ObjCollection As IHTMLElementCollection
    Dim Link As IHTMLElement

    'get link tags on page
    Set ObjCollection = htmlFinancialStatement.getElementsByTagName("a")

    'click on Financials link
    For Each Link In ObjCollection
        If Link.innerHTML = "Financials" Then
            Link.Click
            Exit For
        End If
    Next Link
    
    'Wait until IE is done loading new page
    Do While ie.READYSTATE <> READYSTATE_INTERACTIVE
        Application.StatusBar = "Acquiring financial statement information..."
        DoEvents
    Loop
    
    Do While ie.READYSTATE <> READYSTATE_COMPLETE
        Application.StatusBar = "Acquiring financial statement information..."
        DoEvents
    Loop
    
    'reset status bar
    Application.StatusBar = ""

End Sub
