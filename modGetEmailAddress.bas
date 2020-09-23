Attribute VB_Name = "modGetEmailAddress"
Option Explicit

Public Sub ExtractEmailAddressesFromYahoo()
'************************************
'purpose:extract American Tower Info
'inputs:NONE
'explanation:NONE
'returns:NONE
'************************************
Dim intFreeFile As Integer

   
    '*******************************
    'disable page
    '*******************************
    frmWebBot.MousePointer = vbHourglass
    EnableForm frmWebBot, False
    frmWebBot.sbWebBot.Panels(1).Enabled = True
   
    
    '******************************
    'delete all contents of output file
    '******************************
    intFreeFile = FreeFile
    On Error Resume Next
    Kill App.Path & "\" & strOutputFileName
    On Error GoTo 0
    
    '******************************
    'go through all yahoo pages
    '******************************
Dim lngIndex As Long
Dim blnMorePages As Boolean
Dim strURL As String
Dim strNextUrl As String
    
    'go through all sites
    blnMorePages = True
    lngIndex = 1
    strURL = "http://search.yahoo.com/bin/search?p=Javascript&submit=Search"
    Do While (blnMorePages = True)
        blnMorePages = mblnGetYahooSearchPage( _
            strURL, intFreeFile, lngIndex, strNextUrl)
        strURL = strNextUrl
    Loop
 
    '***********************
    'enable page
    '***********************
    DoEvents
    EnableForm frmWebBot, True
    frmWebBot.MousePointer = vbDefault
    
    frmWebBot.sbWebBot.Panels(1) = ""
    frmWebBot.Refresh

Exit Sub
erhTest:
    DoEvents
    GoTo erhTest
    
End Sub

Private Function mblnGetYahooSearchPage( _
    ByVal vstrUrl As String, _
    ByVal intFreeFile As Integer, _
    ByVal vlngPageNumber As Long, _
    ByVal vstrNextUrl As String) As Boolean
'**********************************
'purpose: get a Yahoo search page
'inputs: vstrUrl--area's url to get
'   intFreeFile--file handle
'   vlngPageNumber--page #
'   vstrNextUrl--URL of next page
'       "" if NONE
'returns:TRUE if there is another page
'        ELSE false
'explanation:
'**********************************
Dim lngIndex As Long

    'set defaults
    mblnGetYahooSearchPage = False
    
    frmWebBot.txtURL = vstrUrl
    frmWebBot.cmGo = True
    
    '**********************************************
    'load page
    '**********************************************
    mLoadPage vstrUrl, vlngPageNumber
    
    '**********************************************
    'put data into database
    '**********************************************
    If (mProcessYahooSearchPage _
        (frmWebBot.txtHtml, intFreeFile, vlngPageNumber) = False) Then
        
        mblnGetYahooSearchPage = False
        Exit Function
    
    End If
    

    '***********************************************
    'look for additional pages
    '***********************************************
    lngPageIndex = rtfPage.Find( _
        "/search?p=Javascript&submit=Search", lngPageIndex)
    If (lngPageIndex = -1) Then
        'no more pages
        
        'set return values
        mblnGetYahooSearchPage = False
    
    Else
        'more pages
    
        'set return values
        mblnGetYahooSearchPage = False
    
    End If
    
    'increment counter
    lngIndex = lngIndex + 1

End Function
