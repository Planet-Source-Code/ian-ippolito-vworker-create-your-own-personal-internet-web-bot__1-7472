Attribute VB_Name = "modYahoo"
 Option Explicit

Public Sub ExtractEmailAddressesFromYahoo( _
    ByVal vstrSearchString As String)
'************************************
'purpose:extracts email address from
'        Yahoo
'inputs:vstrSearchString-string to search on
'   ex: Javascript
'explanation:NONE
'returns:NONE
'************************************
'Dim intFreeFile As Integer
Dim lngPageNumber As Long
Dim blnMorePages As Boolean

Dim strUrl As String
Dim strNextUrl As String

'Dim strHtml As String
Dim rtbHtml As RichTextBox
Dim colUrl As Collection
Dim strLinkURL As Variant

    '*******************************
    'disable page
    '*******************************
    frmWebBot.MousePointer = vbHourglass
    EnableForm frmWebBot, False
    frmWebBot.sbWebBot.Panels(1).Enabled = True
    
    '******************************
    'delete all contents of output file
    '******************************
'    intFreeFile = FreeFile
'    On Error Resume Next
'    Kill App.Path & "\" & strOutputFileName
'    On Error GoTo 0
    
    '******************************
    'go through all yahoo pages
    '******************************

    'init
    blnMorePages = True
    lngPageNumber = 1
    strUrl = "http://search.yahoo.com/bin/search?p=" & vstrSearchString & _
        "&submit=Search"
        
    '***************************
    'go through all search pages
    '***************************
    Do While (blnMorePages = True)
        
        'get next search page
        blnMorePages = mblnGetNextYahooSearchPage( _
            strUrl, vstrSearchString, strNextUrl, rtbHtml) 'strHtml)
        
        'Get all URLs from Yahoo Search Page
        If (mblnProcessYahooSearchPage _
            (rtbHtml, colUrl) = False) Then
            GoTo lblExitLoop
        End If
    
        'crawl URLs for email addresses
        For Each strLinkURL In colUrl
            'frmWebBot.lbUrl.AddItem "Crawling address: " & strLinkURL & " from Yahoo search result page " & lngPageNumber
            frmWebBot.sbWebBot.Panels(1).Text = "Crawling address: " & strLinkURL & " from Yahoo search result page " & lngPageNumber
            Debug.Print "Crawling address: " & strLinkURL & " from Yahoo search result page " & lngPageNumber
            
            Dim blnStayInDomain As Boolean
            
            If (mblnIsYahooUrl(CStr(strLinkURL)) = True) Then
                'yahoo
                blnStayInDomain = False
            Else
                'non yahoo
                blnStayInDomain = True
            End If
            
            'mCrawlUrlForEmailAddresses strLinkURL, 1, 2, _
                blnStayInDomain, _
                3, 1
        Next
    
        'update to next page
        strUrl = strNextUrl
        
        'increment counter
        lngPageNumber = lngPageNumber + 1
        DoEvents
        
lblTest:
DoEvents
'GoTo lblTest

    Loop
 
lblExitLoop:
 
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

Private Function mblnGetNextYahooSearchPage( _
    ByVal vstrURL As String, _
    ByVal vstrSearchString As String, _
    ByRef rstrNextUrl As String, _
    ByRef rrtbText As RichTextBox) As Boolean
'**********************************
'purpose: get a Yahoo search page
'inputs: vstrUrl--url to get
'   vstrSearchString--search string
'        ex. "Javascript"
'   vintFreeFile--file handle
'   vlngPageNumber--page #
'   rstrNextUrl(output)--URL of next page
'                        "" if NONE
'   rrtbText(output)--pointer to rich text
'     box containing HTML
'
'returns:TRUE if there is another page
'        ELSE false
'explanation:
'**********************************
Dim lngIndex As Long

    'set defaults
    mblnGetNextYahooSearchPage = False
    
    '**********************************************
    'load page
    '**********************************************
    frmWebBot.txtURL = vstrURL
    frmWebBot.cmGo = True
    'mLoadPage vstrUrl
    
    '***********************************************
    'look for additional pages
    '***********************************************
Dim lngPageIndex As Long
    lngPageIndex = 1
    lngPageIndex = frmWebBot.txtHtml.Find( _
        "Next", lngPageIndex)
        '"/search?p=" & vstrSearchString & "&submit=Search", lngPageIndex)
    If (lngPageIndex = -1) Then
        'no more pages
        
        'set return values
        mblnGetNextYahooSearchPage = False
    
    Else
        'more pages
    
        'set return values
        mblnGetNextYahooSearchPage = True
        lngPageIndex = lngPageIndex - 100
        lngPageIndex = frmWebBot.txtHtml.Find( _
            "<a href", lngPageIndex)
        rstrNextUrl = "http://search.yahoo.com" & mExtractHTML("""", """", _
            frmWebBot.txtHtml, lngPageIndex)
        
    End If
    
    'set return value
    Set rrtbText = frmWebBot.txtHtml

End Function
Private Function mblnProcessYahooSearchPage( _
    ByVal vrtbHtml As RichTextBox, _
    ByRef rcolUrls As Collection) As Boolean
'**********************************************************
'purpose:process a Yahoo search page
'        by getting all links on page
'inputs:vrtbHtml--rich text box containing information
'       (outputs)rcolUrls--collection of URLs on page
'explanation:expects current page to be on Yahoo search page
'***********************************************************
Dim lngPageIndex As Long
Dim lngPageStartIndex As Long
Dim strLinkURL As String
    
    'init
    Set rcolUrls = Nothing
    Set rcolUrls = New Collection

    'find beginning
    lngPageIndex = 1
    lngPageIndex = vrtbHtml.Find( _
            "Category Match", lngPageIndex)
    lngPageStartIndex = lngPageIndex

    'get next link
    Do While (Left$(strLinkURL, 8) <> "/search?") And _
        (Left$(strLinkURL, 30) <> "http://ink.yahoo.com/bin/query") And _
        lngPageIndex >= lngPageStartIndex And lngPageIndex <> -1

        lngPageIndex = vrtbHtml.Find( _
            "<A HREF=", lngPageIndex)
        strLinkURL = mExtractHTML("""", """", _
            frmWebBot.txtHtml, lngPageIndex)
            
        'check for bad link
        'mailto:
        If InStr(LCase$(strLinkURL), "mailto:") > 0 Then
            GoTo lblNextLoop
        End If
        
        'news:
        If InStr(LCase$(strLinkURL), "news:") > 0 Then
            GoTo lblNextLoop
        End If
        
        'news:
        If InStr(LCase$(strLinkURL), "ftp:") > 0 Then
            GoTo lblNextLoop
        End If
        

        'crawl URL for email addresses
        rcolUrls.Add strLinkURL
lblNextLoop:
    Loop
    
    'set return value
    mblnProcessYahooSearchPage = True
    
End Function
Sub mCrawlUrlForEmailAddresses( _
    ByVal vstrLinkURL As String, _
    ByVal vstrParentLinkUrl As String, _
    ByVal vlngCurrentDepth As Long, _
    ByVal vlngMaxDepth As Long, _
    ByVal vlngMaxEmailsInDomain As Long, _
    ByRef rlngCurrentEmailsInDomain As Long)
'*****************************************************
'purpose:get all email addresses
'inputs:vstrLinkURL--link to start at
'       vstrParentLinkUrl--parent of that link
'       vlngCurrentDepth--current depth (init to 1)
'       vlngMaxDepth--max depth to search to on sub links
'       vlngMaxEmailsInDomain--(when vblnStayInDomain) is
'           true--sets max # of emails to find in domain
'           before leaving it (ex:5 )
'       rlngCurrentEmailsInDomain--used to implement
'           vlngMaxEmailsInDomain--counts # of emails
'           currently gathered in domain (init to 1)
'explanation: yahoo links don't count in depth
'******************************************************
Dim colUrl As Collection
Dim colEmail As Collection
Dim strUrl As Variant
Dim objNode As Node
Dim lngSubjectIndex As Long

    
    '********************************
    'make sure we haven't exceeded max depth
    '********************************
    If vlngCurrentDepth <= vlngMaxDepth _
         Then
    
    
        'set URL as 'being visited'
        mSaveVisitedUrl vstrLinkURL, 1
    
        '********************************
        'send to treeview
        '********************************
        On Error GoTo lblDuplicate
        If vstrParentLinkUrl = "" Then
            Set objNode = frmWebBot.tvLinks.Nodes.Add( _
                , , "page" & vstrLinkURL, vstrLinkURL, "page")
        Else
            Set objNode = frmWebBot.tvLinks.Nodes.Add("page" & vstrParentLinkUrl, _
                tvwChild, "page" & vstrLinkURL, vstrLinkURL, "page")
            objNode.Parent.Expanded = True
        End If
        On Error GoTo 0
        
    
    
        '********************************
        'get all links in current page
        '(note:updates textbox)
        '********************************
        Set colUrl = mcolGetAllUrlsInPage(vstrLinkURL)

         'check for stop on each page
        If (frmWebBot.chkStopAfterEachPage = vbChecked) Then
            frmWebBot.cmReview = True
        End If



        'see if any url's on page
        If colUrl Is Nothing = False Then
           
            '*****************************************
            'make sure newpage is relevant to topic
            '*****************************************
            If (frmWebBot.txtSubject <> "") Then
                lngSubjectIndex = frmWebBot.txtHtml.Find( _
                    frmWebBot.txtSubject, 1)
                If (lngSubjectIndex = -1) Then
                    objNode.Image = "irrelevant"
                    GoTo lblDone
                Else
                    'Stop
                End If
            End If
            
        
        
            '***********************
            'extract email addresses
            '***********************
            Set colEmail = _
                mcolExtractAllEmailAddressesOnPage(vstrLinkURL)
            
            If (colEmail.Count > 0) Then
                mOutputEmailAddresses colEmail, vstrLinkURL
                
                'check if still in domain
                If (strGetDomainName(vstrLinkURL) = _
                    strGetDomainName(vstrParentLinkUrl)) Then
                    rlngCurrentEmailsInDomain = rlngCurrentEmailsInDomain + colEmail.Count
                
                    'check for max emails in domain exceeded
                    If (rlngCurrentEmailsInDomain > vlngMaxEmailsInDomain) Then
                        objNode.Image = "irrelevant"
                        GoTo lblDone
                    End If
                End If
                

            End If
        
        
        
            '********************************
            'go through all links returned
            '********************************
    
            For Each strUrl In colUrl
                
                DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
                
    
                '********************************
                'check if URL should be irrelevant
                '********************************
                If (mblnSkipUrlCrawl(strUrl, vstrLinkURL) = True) Then
                    
                    '********************************
                    'add to treeview
                    '********************************
lblSkip:
                    On Error GoTo lblDuplicate2
                    Set objNode = frmWebBot.tvLinks.Nodes.Add("page" & vstrLinkURL, _
                        tvwChild, "page" & strUrl, strUrl, "skipped")
                    On Error GoTo 0
                    objNode.Parent.Expanded = True
                    GoTo lblNextUrl
                End If
                
                
                '***********************
                'crawl URL recursively
                '***********************
                mCrawlUrlForEmailAddresses strUrl, vstrLinkURL, _
                    vlngCurrentDepth + 1, vlngMaxDepth, _
                    vlngMaxEmailsInDomain, rlngCurrentEmailsInDomain
               
lblNextUrl:
            Next strUrl
    
        End If  'if colURL is nothing = false
        
lblDone:
        'all sub URL's crawled
        'set as 'finish visited'
        mSaveVisitedUrl vstrLinkURL, 2
    
    
    End If
    
Exit Sub
lblDuplicate:
    'done
    Resume lblDone
lblDuplicate2:
    Debug.Print "duplicate found: " & strUrl
    Resume lblNextUrl
End Sub
Function mblnSkipUrlCrawl( _
    ByVal vstrURL As String, _
    ByVal vstrLinkURL As String) As Boolean
'**************************************************
'purpose:determine if URL should be irrelevant
'inputs:vstrURL--url to check
'       vstrLinkURL--last url checked
'explanation:NONE
'returns:TRUE if should be irrelevant, else false
'**************************************************

    'set default
    mblnSkipUrlCrawl = True
    
    'check for same URL as last time
    If (vstrURL = vstrLinkURL) Then
        Exit Function
    End If
    
    'check for yahoo
    If (mblnIsYahooUrl(vstrURL) = True) Then
    
        If frmWebBot.optYahoo = False Then
            'not yahoo search--eject
            Exit Function
        Else
            'yahoo search
            'non directory links are not allowed
            If (InStr(vstrURL, "dir.yahoo.com") = 0) Then
                Exit Function
            End If
            
            'directory link
            'check for link lower in directory
            '(i.e.:
            'http://dir.yahoo.com/Computers_and_Internet/Programming_Languages/JavaScript/Applets/
            'and last link was:
            'http://dir.yahoo.com/Computers_and_Internet/Programming_Languages/JavaScript/
            
'            If (InStr(vstrURL, frmWebBot.FirstYahooSearchURL) > 0) Then
'                'lower link
'                'do nothing
'            Else
'                Exit Function
'            End If
        
        End If
        
    End If
    
'    'check for altavista
'    If (mblnIsAltaVistaUrl(vstrURL) = True) Then
'        Exit Function
'    End If
'
'    'check for hotbot
'    If (mblnIsHotbotUrl(vstrURL) = True) Then
'        Exit Function
'    End If
'
'    'check for infoseek
'    If (mblnIsInfoseekUrl(vstrURL) = True) Then
'        Exit Function
'    End If
'
'    'check for dejanews
'    If (mblnIsDejaNewsUrl(vstrURL) = True) Then
'        Exit Function
'    End If
    
    'check for bookstore
    If (mblnIsBookStore(vstrURL) = True) Then
        Exit Function
    End If
    
    'check for geocities
    If (InStr(LCase$(vstrURL), "geocities.com") > 0) Then
        'help/home
        If (InStr(LCase$(vstrURL), "www.geocities.com/help/") > 0) Then
            Exit Function
        End If
        'home
        If (InStr(LCase$(vstrURL), "www.geocities.com/home/") > 0) Then
            Exit Function
        End If
        'help
        If (InStr(LCase$(vstrURL), "www.geocities.com/help/") > 0) Then
            Exit Function
        End If
        'join
        If (InStr(LCase$(vstrURL), "www.geocities.com/join/") > 0) Then
            Exit Function
        End If
        'membership
        If (InStr(LCase$(vstrURL), "www.geocities.com/members/") > 0) Then
            Exit Function
        End If
    
    End If
    
    'check for already visited
    If (mblnAlreadyVisiting(vstrURL) = True) Then
        'Debug.Print "SKIPPING " & vstrURL
        Exit Function
    End If
    
    'set return value
    mblnSkipUrlCrawl = False
    
End Function

Sub mRemoveExtraneousYahooLinks( _
    ByRef rcolUrls As Collection, _
    ByVal vstrLinkURL As String)
'******************************************
'purpose:remove extraneous Yahoo links
'inputs:rcolUrls--complete list of urls
'   vstrLinkUrl--parent url to compare against
'explanation:
'    header URL (included in final output):
'       http://dir.yahoo.com or non yahoo address
'    footer URL(not included in final output):
'       Yahoo address that doesn't start with:
'       http://dir.yahoo.com
'
'returns:NONE
'******************************************
Dim strUrl As Variant
Dim lngIndex As Long
Dim blnSearchUrlsFound As Boolean
Const strYahooDirectoryHeader = "http://dir.yahoo.com"
Dim colOutputUrls As Collection

    lngIndex = 1
    Set colOutputUrls = New Collection
    blnSearchUrlsFound = False
    For Each strUrl In rcolUrls
        
'Debug.Print strUrl

        'check for link to something higher in yahoo index
        '(which should be ignored)
        'i.e. current directory is:http://dir.yahoo.com/Computers_and_Internet/Programming_Languages/
        'and possible link is:http://dir.yahoo.com/Computers_and_Internet/
        If (InStr(vstrLinkURL, strUrl) > 0) Then
            GoTo lblGetNextUrl
        End If
        
        'remove bad URLs;
        'ex:http://dir.yahoo.com/Business_and_Economy/Companies/Books/Shopping_and_Services/Booksellers/Computers/Titles/Programming_Languages/JavaScript/
        If (InStr(strUrl, "dir.yahoo.com/Business_and_Economy") > 0) Then
            GoTo lblGetNextUrl
        End If
        'ex:news:
        If (InStr(strUrl, "news:") > 0) Then
            GoTo lblGetNextUrl
        End If
        

        'check for header (included in output)
        If (blnSearchUrlsFound = False) Then
            If (InStr(LCase$(strUrl), strYahooDirectoryHeader) > 0) Or _
                (mblnIsYahooUrl(CStr(strUrl)) = False) Then
                
                
                'found header
                blnSearchUrlsFound = True
            End If
        End If

        'check for real audio (remove it)
        If (InStr(strUrl, "http://rd.yahoo.com/") > 0) Then
            GoTo lblGetNextUrl
        End If
        
        If (blnSearchUrlsFound) Then
            'check for footer (not included in output)
            If (mblnIsYahooUrl(CStr(strUrl)) = True) And _
                (InStr(LCase$(strUrl), strYahooDirectoryHeader) = 0) Then
                Exit For
            End If
            
            colOutputUrls.Add strUrl
        End If
        
lblGetNextUrl:
        lngIndex = lngIndex + 1
        
    Next strUrl
    
    'set output
    Set rcolUrls = colOutputUrls
    
End Sub


Sub mOutputEmailAddresses( _
    ByVal vcolEmail As Collection, _
    ByVal vstrURL As String)
'***********************************
'purpose:output email addresses
'inputs:vcolEmail--address to output
'   vstrURL--URL
'***********************************
Dim strEmailAddress As Variant
Dim objConnection As ADODB.Connection
Dim objRecordset As ADODB.Recordset
    

    
    'connect to database
    ConnectToDatabase objConnection

    
    'open recordset
    
    Dim strSQL As String
    strSQL = "SELECT * FROM Extracted_Email_Address " & vbCrLf & _
        "WHERE ExtractedEmailAddressId=-1"
    Set objRecordset = New ADODB.Recordset
    On Error GoTo lblOpenError
    objRecordset.Open strSQL, _
        objConnection, adOpenForwardOnly, _
        adLockPessimistic
    On Error GoTo 0
    
    For Each strEmailAddress In vcolEmail
    

        objRecordset.AddNew
        If Len(strEmailAddress > 50) Then
            strEmailAddress = Left$(strEmailAddress, 50)
        End If
        objRecordset("EmailAddress") = strEmailAddress
        objRecordset("URL") = vstrURL
        On Error GoTo lblUpdate
        objRecordset.Update
        On Error GoTo 0

        'add to listbox
        'frmWebBot.lblEmail.AddItem strEmailAddress
        'frmWebBot.lblEmail.ListIndex = _
            frmWebBot.lblEmail.ListCount - 1
            
        'send to treeview
        Dim objNode As Node
        On Error Resume Next
        Set objNode = frmWebBot.tvLinks.Nodes.Add("page" & vstrURL, _
            tvwChild, "email" & strEmailAddress, strEmailAddress, "email")
        objNode.Parent.Expanded = True
        On Error GoTo 0
            
lblWaitNext:
    Next strEmailAddress
    
    'close recordset
    objRecordset.Close
    Set objRecordset = Nothing
    
    DisconnectFromDatabase objConnection
    
    
Exit Sub
lblDuplicate:
    'duplicate
    objRecordset.CancelUpdate
    On Error Resume Next
    Set objNode = frmWebBot.tvLinks.Nodes.Add("page" & vstrURL, _
        tvwChild, "email" & strEmailAddress & vstrURL, strEmailAddress, "duplicate")
    objNode.Parent.Expanded = True
    On Error GoTo 0
    GoTo lblWaitNext
lblOpenError:
    Err.Raise vbObjectError, "", "Error opening recordset!"
lblUpdate:
    Select Case (Err.Number)
        Case -2147217887
            'duplicate
            Resume lblDuplicate

    End Select
    
End Sub
Sub mSaveVisitedUrl(ByVal vstrURL As String, _
    ByVal vintVisitingStatus As Integer)
'********************************************************
'purpose:set the URL as 'visited'
'inputs:vstrURL--URL to set
'       vintVisitingStatus--visiting status
'       1=visiting(not done), 2=finished visited
'returns:NONE
'explanation:NONE
'********************************************************
Dim objConnection As ADODB.Connection
Dim objRecordset As ADODB.Recordset
    
    'connect to database
    ConnectToDatabase objConnection
    
    'open recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM WebBot_Visited_Url " & vbCrLf & _
        "WHERE url='x'"
    Set objRecordset = New ADODB.Recordset
    On Error GoTo lblOpenError
    objRecordset.Open strSQL, _
        objConnection, adOpenForwardOnly, _
        adLockPessimistic
    On Error GoTo 0
    
    objRecordset.AddNew
    objRecordset("URL") = vstrURL
    objRecordset("VisitStatusId") = vintVisitingStatus
    objRecordset.Update
    
    'close recordset
    objRecordset.Close
    Set objRecordset = Nothing
    
    DisconnectFromDatabase objConnection
Exit Sub
lblOpenError:
End Sub

Sub mDeleteVisitedUrls()
'*******************************************************
'purpose:delete all visited URLS
'inputs:NONE
'returns:NONE
'explanation:NONE
'*******************************************************
Dim objConnection As ADODB.Connection

    If (MsgBox("Are you sure you want to purge all visited URLs?", vbYesNo + vbDefaultButton1, _
        "Purge?") = vbYes) Then
        
        'connect to database
        ConnectToDatabase objConnection
        
        'open recordset
        Dim strSQL As String
        strSQL = "DELETE FROM  WebBot_Visited_Url "
        
        objConnection.Execute strSQL
        
        DisconnectFromDatabase objConnection
        
    End If
    
End Sub
Sub mResetAllVisiting()
'*******************************************************
'purpose:resets all URLs that are set as 'still visiting'
'       to not visited
'inputs:NONE
'returns:NONE
'explanation:NONE
'*******************************************************
Dim objConnection As ADODB.Connection

    'connect to database
    ConnectToDatabase objConnection
    
    'open recordset
    Dim strSQL As String
    strSQL = "UPDATE WebBot_Visited_Url " & vbCrLf & _
        "SET visitstatusid=NULL " & _
        "WHERE visitstatusid=1"
   
    objConnection.Execute strSQL
      
    DisconnectFromDatabase objConnection
    
End Sub
    
Function mblnAlreadyVisiting(ByVal vstrURL As String)
'********************************************************
'purpose:indicates whether URL has been visited or not
'       (note: can be either 'finished visiting' or 'visiting')
'inputs:vstrURL--URL to check
'returns:NONE
'explanation:NONE
'********************************************************
Dim objConnection As ADODB.Connection
Dim objRecordset As ADODB.Recordset
    
    'connect to database
    ConnectToDatabase objConnection
    
    'open recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM WebBot_Visited_Url " & vbCrLf & _
        "WHERE url='" & vstrURL & "'"
    Set objRecordset = New ADODB.Recordset
    On Error GoTo lblOpenError
    objRecordset.Open strSQL, _
        objConnection, adOpenForwardOnly, _
        adLockPessimistic
    On Error GoTo 0
    
    If objRecordset.EOF = False Then
        'found
        mblnAlreadyVisiting = True
    Else
        'not found
        mblnAlreadyVisiting = False
    End If
    
    'close recordset
    objRecordset.Close
    Set objRecordset = Nothing
    
    DisconnectFromDatabase objConnection
    
Exit Function
lblOpenError:

End Function


