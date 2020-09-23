Attribute VB_Name = "AmericanTower"
Option Explicit
Sub mExtractAmericanTowerWebInfo()
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
'    Open App.Path & "\" & strOutputFileName For Output As intFreeFile
'    'close output file
'    Close intFreeFile
    
    '******************************
    'get all sites
    '******************************
Dim lngIndex As Long
Dim boolReturn As Boolean
    
    'go through all sites
    boolReturn = True
    lngIndex = 1
    Do While (boolReturn = True)
        'frmWebBot.sbWebBot.Panels(1) = "Getting site # " & lngIndex
        'frmWebBot.Refresh
        
        boolReturn = mGetSite( _
        "http://www.americantower.com/tower.asp?id=" & lngIndex, _
        intFreeFile, lngIndex)
        
        lngIndex = lngIndex + 1
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

Function mGetSite(ByVal vstrURL As String, _
    ByVal intFreeFile As Integer, _
    ByVal vlngPageNumber As Long) As Boolean
'**********************************
'purpose:get an area
'inputs:vstrUrl--area's url to get
'   intFreeFile--file handle
'   vlngPageNumber--page #
'returns:TRUE on success, else FALSE
'explanation:
'**********************************
Dim lngIndex As Long

    'set defaults
    mGetSite = False
    
    frmWebBot.txtURL = vstrURL
    frmWebBot.cmGo = True
    
    
    '**********************************************
    'load page
    '**********************************************
    'mLoadPage vstrUrl, vlngPageNumber
    
    '**********************************************
    'put data into database
    '**********************************************
    If (mProcessAmericanTowerPage _
        (frmWebBot.txtHtml, intFreeFile, vlngPageNumber) = False) Then
        
        mGetSite = False
        Exit Function
    
    End If
    

    'set return values
    mGetSite = True

    
End Function
Sub mLoadPage2(ByVal vstrURL)
'******************************************
'purpose:load page
'inputs:
'returns:NONE
'******************************************

'    'stall (t1 goes too fast)

'    For lngIndex = 1 To 100000
'        DoEvents
'    Next lngIndex
Dim lngIndex As Long

    frmWebBot.Refresh
lblRetry:
    On Error GoTo erhStillLoading
    'update RTF textbox with page data

    
    gboolFinishedReceived = False
    frmWebBot.Inet1.AccessType = icDirect
    frmWebBot.txtHtml = frmWebBot.Inet1.OpenURL(vstrURL)
    On Error GoTo 0
    
    'check for too fast for server
    If (frmWebBot.txtHtml.Text = "") Then
        Err.Raise 0, "", "no data returned"
    End If
    
Exit Sub

erhStillLoading:
    For lngIndex = 1 To 50000
        DoEvents
    Next lngIndex
    GoTo lblRetry
End Sub
Function mProcessAmericanTowerPage( _
    rtfPage As RichTextBox, _
    vintFreeFile As Integer, _
    ByVal vlngDefaultId As Long) As Boolean
'*****************************
'purpose:process current page
'inputs:rtfPage--rich text box with
'        HTML
'       vintFreeFile--open file
'       default site id if none found
'returns:TRUE if found, Else FALSE
'explanation:NONE
'*****************************
Dim lngPageIndex As Long
Dim lngStringStart, lngStringEnd As Long

Dim strId As String
Dim strLat As String
Dim strLon As String
Dim strState As String
Dim strHeight As String
Dim strDummy As String

    lngPageIndex = 0
    'Do
    
        '********************
        'find site id
        '********************
        
        'extract id
        'find SITE #
        lngPageIndex = 0
        lngPageIndex = rtfPage.Find("SITE NAME", lngPageIndex)
        If (lngPageIndex = -1) Then
            'Exit Do
            mProcessAmericanTowerPage = False
            Exit Function
        End If
        
        lngPageIndex = 0
        lngPageIndex = rtfPage.Find("SITE #", lngPageIndex)
        strId = mExtractHTML("<b>", "</b>", rtfPage, lngPageIndex)
        
        'check for no id
        If (IsNumeric(strId) = False) Then
            strId = vlngDefaultId '"NOT GIVEN"
        End If
            
        'extract State
        lngPageIndex = rtfPage.Find("STATE", lngPageIndex)
        strState = mExtractHTML("<b>", "</b>", rtfPage, lngPageIndex)
            
        'extract Latitude
        lngPageIndex = rtfPage.Find("LATITUDE", lngPageIndex)
        strLat = mExtractHTML("<b>", "&quot;", rtfPage, lngPageIndex)
        
        'extract Longitude
        lngPageIndex = rtfPage.Find("LONGITUDE", lngPageIndex)
        strLon = mExtractHTML("<b>", "&quot;", rtfPage, lngPageIndex)
        
        'skip tower type
        lngPageIndex = rtfPage.Find("SITE TYPE", lngPageIndex)
        If (lngPageIndex > -1) Then
            strDummy = mExtractHTML("<b>", "</b>", rtfPage, lngPageIndex)
        Else
            GoTo lblEndOfPage
        End If
        
        'extract Height
        lngPageIndex = rtfPage.Find("GRND AMSL", lngPageIndex)
        If (lngPageIndex > -1) Then
            strHeight = mExtractHTML("<b>", "</b>", rtfPage, lngPageIndex)
        Else
            GoTo lblEndOfPage
        End If
        
lblEndOfPage:
        'output to file
        'mOutputData strId, strLat, strLon, strState, strHeight, vintFreeFile
        
    'Loop
    
    mProcessAmericanTowerPage = True
    
End Function


