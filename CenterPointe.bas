Attribute VB_Name = "CenterPointe"
Function mExtractCenterPointeWebInfo()
'****************************
'purpose:extract info from web
'inputs:NONE
'returns:NONE
'explanation:NONE
'****************************
Dim lngPage As Long
Dim blnEOF As Boolean
Dim strUrl As String
Const strStartUrl = "http://statusnow.com/cpoint/listings/" & _
        "site2.cfm?sortby1=cpsiteid&submit=Search"


Dim intFreeFile As Integer


    'open output file
    intFreeFile = FreeFile
    Open App.Path & "\" & strOutputFileName For Output As intFreeFile
    
    'close output file
    Close intFreeFile

    'init variables
    blnEOF = False
    strUrl = strStartUrl
    lngPage = 1
    Do
        'mLoadPage strURL, lngPage
        lngPage = lngPage + 1
        
        mProcessCenterPointePage frmWebBot.txtHtml, intFreeFile
        'strUrl = mGetNextUrl(frmWebBot.txtHtml)
        
    Loop While (strUrl <> "")

    
End Function

Sub mProcessCenterPointePage(rtfPage As RichTextBox, _
    vintFreeFile As Integer)
'*****************************
'purpose:process current page
'inputs:rtfPage--rich text box with
'        HTML
'       vintFreeFile--open file
'returns:NONE
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
    Do
    
        '********************
        'find site id
        '********************
        
        'extract id
        'find site3.cfm?
        lngPageIndex = rtfPage.Find("site3.cfm?", lngPageIndex)
        If (lngPageIndex = -1) Then
            Exit Do
        End If
        strId = mExtractHTML(">", "<", rtfPage, lngPageIndex)
        
        
        
        'extract Latitude
        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
        strLat = mExtractHTML(">", "<", rtfPage, lngPageIndex)
        
        'extract Longitude
        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
        strLon = mExtractHTML(">", "<", rtfPage, lngPageIndex)
        
        'Skip city
        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
        strDummy = mExtractHTML(">", "<", rtfPage, lngPageIndex)
        
        'extract State
        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
        strState = mExtractHTML(">", "<", rtfPage, lngPageIndex)
        
        'skip tower type
        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
        strDummy = mExtractHTML(">", "<", rtfPage, lngPageIndex)
        
        'extract Height
        lngPageIndex = rtfPage.Find("size=", lngPageIndex)
        strHeight = mExtractHTML(">", "<", rtfPage, lngPageIndex)
        
        'output to file
        'mOutputData strId, strLat, strLon, strState, strHeight, vintFreeFile
        
    Loop
    
    
End Sub


