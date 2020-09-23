Attribute VB_Name = "Utility"
Option Explicit
Public Const strOutputFileName = "output.txt"
Public gboolFinishedReceived As Boolean
'Public glngPlanetTypeId As Long
 
Public gobjConnection As ADODB.Connection


Function mNavigateToURL(ByRef rIntInternetControl As Inet, _
    ByRef rbrwsBrowserControl As WebBrowser, _
    ByRef rrtfTextBox As RichTextBox, _
    ByRef vstrURL As String _
    ) As Boolean
'*************************************
'purpose:navigate to URL
'inputs:rIntInternetControl--Internet control
'       rbrwsBrowserControl--browser
'       rrtfTextBox--rtf text box
'       vstrURL--URL to navigate to
'*************************************
    
    'set default
    mNavigateToURL = False
    
    On Error GoTo lblOpenError
    rIntInternetControl.URL = vstrURL
    rIntInternetControl.AccessType = icDirect
    
    frmWebBot.sbWebBot.Panels(1).Text = "Loading " & vstrURL & "..."
    rrtfTextBox.Text = rIntInternetControl.OpenURL
    frmWebBot.sbWebBot.Panels(1).Text = ""
    On Error GoTo 0
    
    If (frmWebBot.chkShowInBrowser = vbChecked) Then
        rbrwsBrowserControl.Navigate vstrURL
    End If
    
    'DoEvents: DoEvents: DoEvents: DoEvents
    mNavigateToURL = True
    
Exit Function
lblOpenError:
    Select Case (Err.Number)
        Case 35761
            'timeout
        Case Else
            
    End Select
End Function

Sub EnableForm(frmForm As Form, boolEnable As Boolean)
'enable/disable form

    Dim ctlControl As Control
    For Each ctlControl In frmForm
    
        On Error Resume Next
        If TypeName(ctlControl) = "Menu" Or TypeName(ctlControl) = "Image" Then
            ctlControl.Enabled = True
        Else
            ctlControl.Enabled = boolEnable
        End If
        On Error GoTo 0
        
    Next ctlControl
    
    On Error Resume Next
    frmForm.cmReview.Enabled = True
    frmForm.fraSearchInfo.Enabled = True
    frmWebBot.sbWebBot.Panels(1).Enabled = True
    frmWebBot.tvLinks.Enabled = True
    frmWebBot.mnuVisitedList.Enabled = True
'    frmWebBot.mnuEmailAddresses.Enabled = True
'    frmWebBot.mnuFile.Enabled = True
    frmWebBot.chkStopAfterEachPage.Enabled = True
    frmWebBot.chkShowInBrowser.Enabled = True
    On Error GoTo 0
    
End Sub

Function mReplaceCharacter(strOrigChar, strReplaceChar, strString)
'****************************************************************
' Name: mReplaceCharacter
' Description:Replaces all instances of substring A with sub
'     string B in a string
' By: Ian Ippolito
' Inputs:strString==string to do replacing on
'strOrigChar==orig substring
'strReplaceChar==substring to replace orig substring

' Returns:strString after replacing all instances of strOrigChar with strReplaceChar
' Assumes:None
' Side Effects:None
'****************************************************************
       
'     '**********************************
'     'changes all strOrigChar
'     ' to
'     ' in strString
'     '**********************************
    Dim strResult
    strResult = ""
    '     'traverse string
    Dim intIndex
    For intIndex = 1 To Len(strString)
    
        If (Mid(strString, intIndex, Len(strOrigChar)) = strOrigChar) Then
            '*************
            'match found
            '*************
            strResult = strResult + strReplaceChar
            intIndex = intIndex + Len(strOrigChar) - 1
        Else
            '*************
            'no match
            '*************
             strResult = strResult + Mid(strString, intIndex, 1)
        End If
        
    Next

    mReplaceCharacter = strResult
    
End Function

Function mExtractHTML(ByVal vstrStartDelimiter As String, _
    ByVal vstrEndDelimiter As String, _
    ByRef rrtfHtml As RichTextBox, _
    ByRef rrlngPageIndex As Long) As String
  '**********************************
  'purpose:extract HTML from RTF text box
  'inputs:vstrStartDelimiter --starting delimeter string
  '         (pass in "" to start at current pos)
  '     vstrEndDelimiter-ending delimiter string
  '     rrtfHtml--RTF text box with HTML
  '     (i/o)rrlngPageIndex--current index on page
  'returns:string found (or "" if nothing found)
  '**********************************
  Dim lngStringStart As Long
  Dim lngStringEnd As Long
  On Error GoTo lblError
  
        'find starting delimiter
        If (vstrStartDelimiter <> "") Then
            'normal
            rrlngPageIndex = rrtfHtml.Find(vstrStartDelimiter, rrlngPageIndex + 1)
            lngStringStart = rrlngPageIndex + Len(vstrStartDelimiter)
        Else
            'start at current position
            lngStringStart = rrlngPageIndex
        End If
        
        'find ending delimiter
        rrlngPageIndex = rrtfHtml.Find(vstrEndDelimiter, lngStringStart + 1)
        lngStringEnd = rrlngPageIndex - 1
        
        'extract text
        rrtfHtml.SelStart = lngStringStart
        rrtfHtml.SelLength = lngStringEnd - lngStringStart + 1
        mExtractHTML = rrtfHtml.SelText
        
        'set output value
        rrlngPageIndex = lngStringEnd + Len(vstrEndDelimiter)
        
On Error GoTo 0

Exit Function
lblError:
    mExtractHTML = "ERROR"
End Function

Function mcolExtractAllEmailAddressesOnPage( _
    ByVal vstrURL As String) As Collection
'***************************************************
'purpose:extract all email addresses on page
'inputs:vstrUrl--url to load from
'returns:collection of email addresses
'explanation:looking for:
'   <a href="mailto:webmaster@humor.com">
'   <a href=mailto:webmaster@humor.com>
'***************************************************
Dim colUrl As Collection

Dim lngPageIndex As Long
Dim lngEndOfUrl As Long
Dim lngPageStartIndex As Long
Dim strLinkURL As String
    
    'init
    Set colUrl = New Collection

    '**********************************************
    'load page
    '**********************************************
    'frmWebBot.txtURL = vstrURL
    'frmWebBot.cmGo = True
    
    'find beginning
    lngPageIndex = 1
    lngPageIndex = frmWebBot.txtHtml.Find( _
            "mailto:", lngPageIndex)

    If (lngPageIndex = -1) Then
        'not found
        GoTo lblExit
    Else
        'found
        lngPageStartIndex = lngPageIndex
    End If
    
    'get next email address
    Do While (lngPageIndex >= lngPageStartIndex)

        'extract email address
        strLinkURL = mExtractHTML(":", ">", _
            frmWebBot.txtHtml, lngPageIndex)
        If (lngPageIndex < lngPageStartIndex) Then
            Exit Do
        End If
        
        'remove quotes (if any)
        strLinkURL = mReplaceCharacter("""", "", strLinkURL)
        strLinkURL = mReplaceCharacter("'", "", strLinkURL)
        
        'check for followed by parms:
        'ex. me@place.com?subject=stuff
        If InStr(strLinkURL, "?") > 0 Then
            strLinkURL = Left$(strLinkURL, InStr(strLinkURL, "?") - 1)
        End If
        
        'check for 'bad' anti-spam email addresses
        '<a href="mailto:abuse@concentric.net"></a>
        '<a href="mailto:fraud@uu.net"></a>
        '<a href="mailto:abuse@aol.com"></a>
        '<a href="mailto:abuse@compuserve.com"></a>
        '<a href="mailto:abuse@netcom.com"></a>
        '<a href="mailto:root@ftc.gov"></a>
        '<a href="mailto:uce@ftc.gov"></a>
        '<a href="mailto:root@fcc.gov"></a>
        '<a href="mailto:root@[127.0.0.1]"></a>
        
        If (InStr(LCase$(strLinkURL), "127.0.0.1") > 0) Then
            GoTo lblGetNext
        End If
        If (InStr(LCase$(strLinkURL), "abuse") > 0) Then
            GoTo lblGetNext
        End If
        If (InStr(LCase$(strLinkURL), "fraud") > 0) Then
            GoTo lblGetNext
        End If
        If (InStr(LCase$(strLinkURL), "ftc.gov") > 0) Then
            GoTo lblGetNext
        End If
        If (InStr(LCase$(strLinkURL), "root") > 0) Then
            GoTo lblGetNext
        End If
        
        'other common dead ends
        If (InStr(LCase$(strLinkURL), "subscribe") > 0) Then
            GoTo lblGetNext
        End If
        If (InStr(LCase$(strLinkURL), "support") > 0) Then
            GoTo lblGetNext
        End If
        If (InStr(LCase$(strLinkURL), "sales") > 0) Then
            GoTo lblGetNext
        End If
        If (InStr(LCase$(strLinkURL), "info") > 0) Then
            GoTo lblGetNext
        End If
        If (InStr(LCase$(strLinkURL), "request") > 0) Then
            GoTo lblGetNext
        End If
        
        'remove target: x@x.com target=frame
        If (InStr(LCase$(strLinkURL), " target") > 0) Then
            strLinkURL = Left$(strLinkURL, InStr(LCase$(strLinkURL), " target") - 1)
        End If
        
        'add email to collection
        colUrl.Add strLinkURL

lblGetNext:

        'get next link
        lngPageIndex = frmWebBot.txtHtml.Find( _
            "mailto", lngPageIndex)


    Loop

lblExit:

    'set return value
    Set mcolExtractAllEmailAddressesOnPage = colUrl
    
End Function

Function mcolGetAllUrlsInPage( _
    ByVal vstrURL As String) As Collection
'**************************************
'pursose:get all URLs on page specified
'inputs:vstrURL--URL to check
'returns:collection of URLs on page
'explanation:ssumes href is first in tag
'   (doesn't yet handle <a target=xxx href=xxx>
'   for example)
'side effects:changes frmWebBot.txtHtml
'          and other form controls
'**************************************
Dim colUrl As Collection
Dim lngURLCount As Long
Dim lngPageIndex As Long
Dim lngEndOfUrl As Long
Dim lngPageStartIndex As Long
Dim strLinkURL As String
    
    'init
    Set colUrl = New Collection

    '**********************************************
    'load page
    '**********************************************
    frmWebBot.txtURL = vstrURL
    frmWebBot.cmGo = True
    
    
    'check for not found
    If (InStr(LCase$(frmWebBot.txtHtml.Text), "404 not found") > 0) Then
        Exit Function
    End If
    
    If (InStr(LCase$(frmWebBot.txtHtml.Text), "file not found") > 0) Then
        Exit Function
    End If
    
    If (InStr(LCase$(frmWebBot.txtHtml.Text), "page cannot be found") > 0) Then
        Exit Function
    End If
    
    'find beginning
    lngURLCount = 0
    lngPageIndex = 1
    lngPageIndex = frmWebBot.txtHtml.Find( _
            "<A HREF", lngPageIndex)

    If (lngPageIndex = -1) Then
        'not found
        GoTo lblExit
    Else
        'found
        lngPageStartIndex = lngPageIndex
    End If
    
    'get next link
    Do While (lngPageIndex >= lngPageStartIndex)

        lngPageStartIndex = lngPageIndex
        frmWebBot.sbWebBot.Panels(1).Text = "Searching current page at index: " & lngPageStartIndex
        
        'extract link
        strLinkURL = mExtractHTML("=", ">", _
            frmWebBot.txtHtml, lngPageIndex)
'Debug.Print strLinkURL

        strLinkURL = mstrFormatAndValidateUrl(strLinkURL, vstrURL)
        
        'add URL to collection
        If strLinkURL <> "" Then
            On Error GoTo lblDuplicate
            lngURLCount = lngURLCount + 1
            colUrl.Add strLinkURL, "KEY" & strLinkURL
            On Error GoTo 0
        End If
        'get next link
lblGetNext:
        lngPageIndex = frmWebBot.txtHtml.Find( _
            "<A HREF", lngPageIndex)

        If (lngURLCount Mod 10 = 0) Then
            DoEvents
        End If
lblNext:
    Loop

lblExit:
    frmWebBot.sbWebBot.Panels(1).Text = ""
    'set return value
    Set mcolGetAllUrlsInPage = colUrl
Exit Function
lblDuplicate:
    Resume lblGetNext
End Function
Function strGetDirectory(ByVal vstrURL As String) As String
'***********************************
'purpose:get fully qualified directory name from fully qualified URL
'inputs:vstrURL -url to extract from
'returns:dir name or "" if not found
' note:includes /
'explanation:NONE
'example: http://www.yahoo.com/xxx/ddd/rrr/eee.htm ->
'         http://www.yahoo.com/xxx/ddd/rrr/
'***********************************
Dim lngIndex As Long

    lngIndex = InStrRev(vstrURL, "/")
    If (lngIndex > 0) Then
        strGetDirectory = Left$(vstrURL, lngIndex)
    Else
        strGetDirectory = ""
    End If


End Function
Function strGetDomainName(ByVal vstrURL As String) As String
'***********************************
'purpose:get domain name from fully qualified URL
'inputs:vstrURL -url to extract from
'returns:domain name or "" if not found
' note:includes /
'explanation:NONE
'example: http://www.yahoo.com/xxx/ddd/rrr/eee.htm ->
'         http://www.yahoo.com/
'***********************************
Dim lngIndex As Long

    lngIndex = InStr(vstrURL, "http://")
    If (lngIndex > 0) Then
        lngIndex = lngIndex + Len("http://")
        lngIndex = InStr(lngIndex, vstrURL, "/")
        If (lngIndex > 0) Then
            strGetDomainName = Left$(vstrURL, lngIndex)
            Exit Function
        End If
    End If

    strGetDomainName = ""
    
End Function



Function mblnIsAltaVistaUrl( _
    vstrLinkURL As String) As Boolean
'********************************************
'purpose:determines if a URL is a yahoo URL
' i.e. contains yahoo.com
'inputs:vstrLinkURL-url to test
'returns:NONE
'explanation:NONE
'********************************************
    If (InStr(LCase$(vstrLinkURL), "altavista.com") > 0) Then
        mblnIsAltaVistaUrl = True
    Else
        mblnIsAltaVistaUrl = False
    End If
End Function
Function mblnIsHotbotUrl( _
    vstrLinkURL As String) As Boolean
'********************************************
'purpose:determines if a URL is a yahoo URL
' i.e. contains yahoo.com
'inputs:vstrLinkURL-url to test
'returns:NONE
'explanation:NONE
'********************************************
    If (InStr(LCase$(vstrLinkURL), "hotbot.com") > 0) Then
        mblnIsHotbotUrl = True
    Else
        mblnIsHotbotUrl = False
    End If
    
End Function

Function mblnIsInfoseekUrl( _
    vstrLinkURL As String) As Boolean
'********************************************
'purpose:determines if a URL is a yahoo URL
' i.e. contains yahoo.com
'inputs:vstrLinkURL-url to test
'returns:NONE
'explanation:NONE
'********************************************
    
    If (InStr(LCase$(vstrLinkURL), "hotbot.com") > 0) Then
        mblnIsInfoseekUrl = True
    Else
        mblnIsInfoseekUrl = False
    End If
    
End Function
Function mblnIsDejaNewsUrl( _
    vstrLinkURL As String) As Boolean
'********************************************
'purpose:determines if a URL is a yahoo URL
' i.e. contains yahoo.com
'inputs:vstrLinkURL-url to test
'returns:NONE
'explanation:NONE
'********************************************
    
    If (InStr(LCase$(vstrLinkURL), "hotbot.com") > 0) Then
        mblnIsDejaNewsUrl = True
    Else
        mblnIsDejaNewsUrl = False
    End If
    
End Function
Function mblnIsBookStore( _
    vstrLinkURL As String) As Boolean
'********************************************
'purpose:determines if a URL is a yahoo URL
' i.e. contains yahoo.com
'inputs:vstrLinkURL-url to test
'returns:NONE
'explanation:NONE
'********************************************
    
    If (InStr(LCase$(vstrLinkURL), "fatbrain.com") > 0) Then
        mblnIsBookStore = True
    ElseIf (InStr(LCase$(vstrLinkURL), "amazon.com") > 0) Then
        mblnIsBookStore = True
    ElseIf (InStr(LCase$(vstrLinkURL), "barnesandnoble.com") > 0) Then
        mblnIsBookStore = True
    Else
        mblnIsBookStore = False
    End If
    
End Function

Function mblnIsYahooUrl( _
    vstrLinkURL As String) As Boolean
'********************************************
'purpose:determines if a URL is a yahoo URL
' i.e. contains yahoo.com
'inputs:vstrLinkURL-url to test
'returns:NONE
'explanation:NONE
'********************************************
    
    If (InStr(LCase$(vstrLinkURL), "yahoo.com") > 0) Then
        mblnIsYahooUrl = True
    Else
        mblnIsYahooUrl = False
    End If
    
End Function
Function mblnIsYahooIndexUrl( _
    vstrLinkURL As String) As Boolean
'********************************************
'purpose:determines if a URL is a yahoo INDEX URL
' (i.e. contains dir.yahoo.com
'inputs:vstrLinkURL-url to test
'returns:NONE
'explanation:NONE
'********************************************
    
    If (InStr(LCase$(vstrLinkURL), "dir.yahoo.com") > 0) Then
        mblnIsYahooIndexUrl = True
    Else
        mblnIsYahooIndexUrl = False
    End If
    
End Function


'Function mExtractInformationFromWeb()
''****************************
''purpose:extract info from web
''inputs:NONE
''returns:NONE
''explanation:NONE
''****************************
'Dim lngPage As Long
'Dim blnEOF As Boolean
'Dim strUrl As String
'Const strStartUrl = "http://statusnow.com/cpoint/listings/" & _
'        "site2.cfm?sortby1=cpsiteid&submit=Search"
'
'
'Dim intFreeFile As Integer
'
'
'    'open output file
'    intFreeFile = FreeFile
'    Open App.Path & "\" & strOutputFileName For Output As intFreeFile
'
'    'close output file
'    Close intFreeFile
'
'    'init variables
'    blnEOF = False
'    strUrl = strStartUrl
'    lngPage = 1
'    Do
'        'mLoadPage strURL, lngPage
'        lngPage = lngPage + 1
'
'        mProcessPage frmWebBot.txtHtml, intFreeFile
'        strUrl = mGetNextUrl(frmWebBot.txtHtml)
'
'    Loop While (strUrl <> "")
'
'
'End Function
'
'
'Sub mLoadPage(ByVal vstrUrl As String)
''********************************************
''purpose:load a page into 2 controls
''inputs:vstURl--url to load
''       vlngPageCount--current page #
''returns:NONE
''explanation:none
''********************************************
'
'    'disable page
'    frmWebBot.MousePointer = vbHourglass
'    EnableForm frmWebBot, False
'    frmWebBot.sbWebBot.Panels(1).Enabled = True
'    'frmWebBot.sbWebBot.Panels(1) = "Loading page " & Trim$(vlngPageCount) & "...please wait"
'
'
'    'update RTF textbox with page data
'    frmWebBot.Inet1.URL = vstrUrl
'    'Inet1.Password = "I(3Lei#4"
'    'Inet1.UserName = "Jonne Smythe"
'
'    'stall (t1 goes too fast)
'    Dim lngIndex As Long
'    For lngIndex = 1 To 100000
'        DoEvents
'    Next lngIndex
'
'    frmWebBot.Refresh
'    On Error GoTo erhStillLoading
'retry:
'    frmWebBot.txtHtml.Text = frmWebBot.Inet1.OpenURL
'
'    'check for too fast for server
'    If (frmWebBot.txtHtml.Text = "") Then
'        Err.Raise 0, "", "no data returned"
'    End If
'    frmWebBot.Refresh
'
'    On Error GoTo 0
'
'    DoEvents
'    frmWebBot.Refresh
'
'    'load browser on page
'    'frmWebBot.brwsWebBot.Navigate vstrUrl
'
'    DoEvents
'    frmWebBot.Refresh
'
'    'enable page
'    EnableForm frmWebBot, True
'    frmWebBot.MousePointer = vbDefault
'    frmWebBot.sbWebBot.Panels(1) = ""
'    frmWebBot.Refresh
'
'Exit Sub
'
'erhStillLoading:
'    For lngIndex = 1 To 50000
'        DoEvents
'    Next lngIndex
'    GoTo retry
'End Sub


'Sub mProcessPage(rtfPage As RichTextBox, _
'    vintFreeFile As Integer)
''*****************************
''purpose:process current page
''inputs:rtfPage--rich text box with
''        HTML
''       vintFreeFile--open file
''returns:NONE
''explanation:NONE
''*****************************
'Dim lngPageIndex As Long
'Dim lngStringStart, lngStringEnd As Long
'
'Dim strId As String
'Dim strLat As String
'Dim strLon As String
'Dim strState As String
'Dim strHeight As String
'Dim strDummy As String
'
'    lngPageIndex = 0
'    Do
'
'        '********************
'        'find site id
'        '********************
'
'        'extract id
'        'find site3.cfm?
'        lngPageIndex = rtfPage.Find("site3.cfm?", lngPageIndex)
'        If (lngPageIndex = -1) Then
'            Exit Do
'        End If
'        strId = mExtractHTML(">", "<", rtfPage, lngPageIndex)
'
'
'
'        'extract Latitude
'        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
'        strLat = mExtractHTML(">", "<", rtfPage, lngPageIndex)
'
'        'extract Longitude
'        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
'        strLon = mExtractHTML(">", "<", rtfPage, lngPageIndex)
'
'        'Skip city
'        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
'        strDummy = mExtractHTML(">", "<", rtfPage, lngPageIndex)
'
'        'extract State
'        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
'        strState = mExtractHTML(">", "<", rtfPage, lngPageIndex)
'
'        'skip tower type
'        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
'        strDummy = mExtractHTML(">", "<", rtfPage, lngPageIndex)
'
'        'extract Height
'        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
'        strHeight = mExtractHTML(">", "<", rtfPage, lngPageIndex)
'
'        'output to file
'        mOutputData strId, strLat, strLon, strState, strHeight, vintFreeFile
'
'    Loop
'
'
'End Sub
        


'
'Function mGetNextUrl(rrtfHtml As RichTextBox) As String
''****************************
''purpose:get next URL from HTML
''inputs:rtfPage--rtf text box with HTML in it
''returns:URL or "" if none
''explanation:NONE
''****************************
'Dim strStart As String
'Dim lngPageIndex As Long
'
'Dim blnNextFiveFound As Boolean
'Dim blnPreviousFiveFound As Boolean
'
'    '******************************************************************
'    'determine what buttons are on page--can be either one or both of:
'    'previous 5 and next 5
'    '******************************************************************
'
'    'verify that this is parm for 'next' and not 'previous'
'    lngPageIndex = rrtfHtml.Find("Previous 5", 1)
'    If (lngPageIndex <> -1) Then
'        blnPreviousFiveFound = True
'    Else
'        blnPreviousFiveFound = False
'    End If
'
'    lngPageIndex = rrtfHtml.Find("Next 5", 1)
'    If (lngPageIndex <> -1) Then
'        blnNextFiveFound = True
'    Else
'        blnNextFiveFound = False
'    End If
'
'
'    'check for no next button
'    If (blnNextFiveFound = False) Then
'        mGetNextUrl = ""
'        Exit Function
'    End If
'
'    'get next button (either 1st or 2nd depending on if other
'    'button exists
'Dim intCount As Integer
'Dim intIndex As Integer
'    If (blnPreviousFiveFound = True) Then
'        intCount = 2
'    Else
'        intCount = 1
'    End If
'
'
'    lngPageIndex = 1
'    For intIndex = 1 To intCount
'        lngPageIndex = rrtfHtml.Find("<form", lngPageIndex)
'        lngPageIndex = rrtfHtml.Find("start", lngPageIndex)
'        strStart = mExtractHTML("value=" & Chr(34), Chr(34), rrtfHtml, lngPageIndex)
'
'    Next
'
'    'set return value
'    mGetNextUrl = "http://statusnow.com/cpoint/listings/" & _
'            "site2.cfm?sortby1=cpsiteid&submit=Search&start=" & Trim(strStart)
'
'End Function
'Sub mOutputData(ByVal vstrId As String, _
'    ByVal vstrLat As String, _
'    ByVal vstrLon As String, _
'    ByVal vstrState As String, _
'    ByVal vstrHeight As String, _
'    ByVal vintFreeFile As Integer)
''****************************
''purpose:output line info
''****************************
'
'    Open App.Path & "\" & strOutputFileName For Append As vintFreeFile
'
'    Write #vintFreeFile, vstrId, vstrLat, vstrLon, vstrState, vstrHeight
'
'    Close #vintFreeFile
'End Sub
Sub mEnumerateCollection(ByVal vcolCollection As Collection)
'*********************************
'purpose:enumerates a collection
'inputs:vcolCollection-collection to enumerate
'explanation:NONE
'returns:NONE
'*********************************
Dim objVariant As Variant

    For Each objVariant In vcolCollection
        Debug.Print objVariant
    Next objVariant

End Sub
Function mstrFormatAndValidateUrl(ByVal vstrLinkURL As String, _
    ByVal vstrParentURL As String) As String
'**************************************************
'purpose:format URL and validate it
'inputs:vstrLinkURL--url to look at
'       vstrParentURL--parent
'returns:URL formatted ("" if invalid
'**************************************************
Dim strLinkURL

    strLinkURL = vstrLinkURL
    
    '************************
    'look for invalid url
    '************************
    'mailto:
    If InStr(LCase$(strLinkURL), "mailto:") > 0 Then
        mstrFormatAndValidateUrl = ""
        Exit Function
        'GoTo lblNext
    End If
    
    'news:
    If InStr(LCase$(strLinkURL), "news:") > 0 Then
        mstrFormatAndValidateUrl = ""
        Exit Function
    End If
    
    'news:
    If InStr(LCase$(strLinkURL), "ftp:") > 0 Then
        mstrFormatAndValidateUrl = ""
        Exit Function
    End If
    
    '***************************
    'remove quotes (if any)
    '***************************
    'preceeding double quotes
    If (Left$(strLinkURL, 1) = Chr(34)) Then
        strLinkURL = Mid$(strLinkURL, 2)
    End If
    'preceeding single quotes
    strLinkURL = mReplaceCharacter("""", "", strLinkURL)
    strLinkURL = mReplaceCharacter("'", "", strLinkURL)
    
        
    'check for additional tags
    'ex:http://www.geocities.com/" target="_top
    If (InStr(strLinkURL, " ") > 0) Then
        strLinkURL = Left$(strLinkURL, InStr(strLinkURL, " "))
    End If
    
    'check for bookmarks in current page
    'ex:http://www.nmt.edu/tcc/help.htm#test

    If (InStr(strLinkURL, "#") > 0) Then
        'check if first . is before last /
         strLinkURL = Left$(strLinkURL, InStr(strLinkURL, "#") - 1)
    End If
            
    'check for malformed URL with "phantom dirs"
    'ex:http://www.nmt.edu/tcc/help/lang/cfamily.html/fortran/fortran/homepage.html
'Dim lngSlash As Long
'        If (InStr(strLinkURL, ".") > 0) Then
'            'check if first . is before last /
'            If (InStr(strLinkURL, ".") < InStrRev(strLinkURL, "/")) Then
'                'cut off at slash after .
'                lngSlash = InStr(InStr(strLinkURL, "."), strLinkURL, "/")
'                strLinkURL = Left$(strLinkURL, lngSlash - 1)
'            End If
'        End If
    
    '3 cases
    '1) http://www.xxx.com/dir/   -> no conversion
    '2) /somedir/default.asp      -> http://www.xx.com/somedir/default.asp
    '3) somedir/default.asp       -> http://www.xx.com/dir/somedir/default.asp
    'fully qualify relative links
    
    'check for case 3
    If (Left$(strLinkURL, 1) <> "/") And (Left$(strLinkURL, 4) <> "http") Then
        
        'prepend current directory and slash
        strLinkURL = strGetDirectory(vstrParentURL) & strLinkURL
        
        'GoTo lblPrePendDomain
'            If (Right$(vstrParentURL, 1) <> "/") Then
'                'URL doesn't end in /
'                strLinkURL = vstrParentURL & strLinkURL
'            Else
'                'URL ends in /
'                strLinkURL = Left$(vstrParentURL, Len(vstrParentURL) - 1) & strLinkURL
'            End If
    ElseIf (Left$(strLinkURL, 1) = "/") Then
        'case 2
        'relative link
lblPrePendDomain:
        'prepend URL
        Dim strDomain As String
        '(remove ending slash from URL)
        If (strGetDomainName(vstrParentURL) <> "") Then
            strDomain = Left$(strGetDomainName(vstrParentURL), Len(strGetDomainName(vstrParentURL)) - 1)
            strLinkURL = strDomain & strLinkURL
        Else
            strDomain = vstrParentURL
            strLinkURL = strDomain
        End If
        
        'If (Right$(vstrParentURL, 1) <> "/") Then
            'URL doesn't end in /
            'strLinkURL = strDomain & strLinkURL
        'Else
            'URL ends in /
            'strLinkURL = strDomain & strLinkURL
        'End If
    End If
    
    mstrFormatAndValidateUrl = strLinkURL
    
End Function

Public Function mintSelectedControlId(ByRef roptControl As Variant)
'***************************************
'purpose: get id of control picked
'   from option button array
'inputs:roptControl--control array to check
'
'***************************************
Dim lngIndex As Long
    
    For lngIndex = roptControl.LBound To roptControl.UBound
        If roptControl(lngIndex) = True Then
            mintSelectedControlId = lngIndex
            Exit Function
        End If
    Next
    
    'not found
    mintSelectedControlId = -1

End Function


