VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmWebBot 
   Caption         =   "Research Web Bot"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8175
   Icon            =   "frmWebBot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Key"
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   5040
      Width           =   8130
      Begin VB.Image Image5 
         Height          =   480
         Left            =   6120
         Picture         =   "frmWebBot.frx":0442
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "Skipped Email"
         Height          =   375
         Left            =   6720
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Root Page"
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   120
         Picture         =   "frmWebBot.frx":0884
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Normal Page"
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1560
         Picture         =   "frmWebBot.frx":0CC6
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Email Address"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4440
         Picture         =   "frmWebBot.frx":1108
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Irrelevant"
         Height          =   252
         Left            =   3720
         TabIndex        =   14
         Top             =   360
         Width           =   732
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3120
         Picture         =   "frmWebBot.frx":154A
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CommandButton cmJava 
      Caption         =   "Yahoo Javascript"
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton cmC 
      Caption         =   "Yahoo C/C++"
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Top             =   8520
      Width           =   1455
   End
   Begin ComctlLib.TreeView tvLinks 
      Height          =   3510
      Left            =   0
      TabIndex        =   10
      Top             =   1545
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6191
      _Version        =   327682
      Indentation     =   471
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1200
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmGo 
      Caption         =   "&Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Frame fraSearchInfo 
      Caption         =   "Search Info:"
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8160
      Begin VB.Frame fraStart 
         Caption         =   "&Start"
         Height          =   1335
         Left            =   2640
         TabIndex        =   24
         Top             =   120
         Width           =   3975
         Begin VB.OptionButton optInfoseek 
            Caption         =   "&Infoseek"
            Height          =   255
            Left            =   1440
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optAltaVista 
            Caption         =   "&Alta Vista"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optCustomURL 
            Caption         =   "&Custom URL:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optHotBot 
            Caption         =   "&Hotbot"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optYahoo 
            Caption         =   "&Yahoo"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox txtSearchURL 
            Height          =   285
            Left            =   1440
            TabIndex        =   25
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Text            =   "Neural Network"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkStopAfterEachPage 
         Caption         =   "&Stop each page"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtSearchDepth 
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Text            =   "3"
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkShowInBrowser 
         Caption         =   "Show in &Browser"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmStartSearch 
         Caption         =   "Start Search"
         Height          =   375
         Left            =   6720
         TabIndex        =   8
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton cmReview 
         Caption         =   "&Pause"
         Height          =   375
         Left            =   6720
         TabIndex        =   7
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label8 
         Caption         =   "Subject:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Search Depth:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
   End
   Begin ComctlLib.StatusBar sbWebBot 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   5910
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   21167
            MinWidth        =   21167
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwsWebBot 
      Height          =   3480
      Left            =   3915
      TabIndex        =   0
      Top             =   1575
      Width           =   4245
      ExtentX         =   7488
      ExtentY         =   6138
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin RichTextLib.RichTextBox txtHtml 
      Height          =   1410
      Left            =   2760
      TabIndex        =   1
      Top             =   7080
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2487
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmWebBot.frx":198C
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWebBot.frx":1A81
            Key             =   "rootpage"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWebBot.frx":1D9B
            Key             =   "skipped"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWebBot.frx":20B5
            Key             =   "duplicate"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWebBot.frx":23CF
            Key             =   "irrelevant"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWebBot.frx":26E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWebBot.frx":2A03
            Key             =   "page"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWebBot.frx":2D1D
            Key             =   "email"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   9000
      Width           =   4305
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEmailAddresses 
      Caption         =   "&Email Addresses"
      Visible         =   0   'False
      Begin VB.Menu mnuReview 
         Caption         =   "&Review Email Addresses"
      End
   End
   Begin VB.Menu mnuVisitedList 
      Caption         =   "&Visited List"
      Begin VB.Menu mnuShowVisitedList 
         Caption         =   "Show 'visited' list"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSeperator2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPurgeVisitedList 
         Caption         =   "&Purge 'visited' list"
      End
   End
End
Attribute VB_Name = "frmWebBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReview As Boolean
Dim mstrStartURL As String
Private mstrFirstYahooSearchURL As String

Public Property Get FirstYahooSearchURL() As String
    FirstYahooSearchURL = mstrFirstYahooSearchURL
End Property
 
Public Property Get StartURL() As String
    StartURL = mstrStartURL
End Property
Private Sub brwsWebBot_DownloadBegin()
    lblStatus = "Download begin..."
End Sub

Private Sub brwsWebBot_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    lblStatus = "Downloaded " & Progress & " out of " & _
        ProgressMax
End Sub

Private Sub brwsWebBot_StatusTextChange(ByVal Text As String)
    lblStatus = "Status change " & Text
End Sub

Private Sub cmAmericanTowern_Click()

End Sub

'Private Sub cmAmericanTower_Click()
'    mExtractAmericanTowerWebInfo
'End Sub

Private Sub cmCancel_Click()
    Unload Me
End Sub

Private Sub cmC_Click()
    'glngPlanetTypeId = 3
    ExtractEmailAddressesFromYahoo "C%2B%2B" '"C++"
End Sub

'Private Sub cmOk_Click()
'    mExtractCenterPointeWebInfo
'End Sub


Private Sub cmGo_Click()
    
    mNavigateToURL Inet1, brwsWebBot, txtHtml, txtURL

End Sub




Private Sub cmJava_Click()
    'glngPlanetTypeId = 2
    ExtractEmailAddressesFromYahoo "Javascript"
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmReview_Click()
    If (mblnReview = False) Then
        'first time
        mblnReview = True
        cmReview.Caption = "&Continue"
        'EnableForm Me, True
        MousePointer = vbNormal
        Do While (mblnReview = True)
            DoEvents
        Loop
        MousePointer = vbHourglass
    Else
        'second time
        mblnReview = False
        cmReview.Caption = "&Pause"
        'EnableForm Me, False
    End If
End Sub

Private Sub cmStartSearch_Click()
Dim objConnection As ADODB.Connection

    'check if new search
    If (MsgBox("Is this a brand new search?", vbYesNo, "New Search?") = vbYes) Then
    
        'purge visited URLs
        
         'connect to database
        ConnectToDatabase objConnection
        
        'open recordset
        Dim strSQL As String
        strSQL = "DELETE FROM  WebBot_Visited_Url "
        
        objConnection.Execute strSQL
        
        DisconnectFromDatabase objConnection
    End If
    
    'get starting URL
    mstrStartURL = mstrGetStartingUrl(Me)
    
    'do search
    SearchGeneric mstrStartURL
    
    'inform user that search is done
    MsgBox "Search Complete", vbOKOnly, "Search Complete"
    
End Sub

Function mstrGetStartingUrl(ByRef rfrmWebBot As frmWebBot) As String
'*****************************************************
'purpose:get starting URL
'inputs:rfrmWebBot--form
'returns:see purpose
'explanation:NONE
'*****************************************************

    'look for custom
    If (optCustomURL = True) Then
        mstrGetStartingUrl = rfrmWebBot.txtSearchURL
        Exit Function
    End If
    
    'look for search
    If (optYahoo = True) Then
        mstrGetStartingUrl = "http://search.yahoo.com/bin/search?p=" & _
            mstrConvertToHtml(rfrmWebBot.txtSubject) & _
            "&submit=Search"
    End If
    If (optHotBot = True) Then
        mstrGetStartingUrl = "http://www.hotbot.com/?MT=" & _
            mstrConvertToHtml(rfrmWebBot.txtSubject) & _
            "&SM=MC&DV=0&LG=any&DC=10&DE=2&BT=H"
    End If
    
    If (optAltaVista = True) Then
        mstrGetStartingUrl = "http://www.altavista.com/cgi-bin/query?pg=q&q=" & _
            mstrConvertToHtml(rfrmWebBot.txtSubject) & _
            "&kl=XX&stype=stext"

    End If

    If (optInfoseek = True) Then
        mstrGetStartingUrl = "http://infoseek.go.com/Titles?qt=" & _
            mstrConvertToHtml(rfrmWebBot.txtSubject) & _
            "&col=WW&sv=IS&lk=noframes&svx=home_searchbox"

    End If
    
End Function

Function mstrConvertToHtml(ByVal vstrText As String) As String
'****************************************
'purpose:convert text to HTML
'inputs:
'returns:
'explanation:
'****************************************
    
    mstrConvertToHtml = mReplaceCharacter("+", "%2B", vstrText)
    mstrConvertToHtml = mReplaceCharacter(" ", "+", vstrText)

End Function



Sub SearchGeneric(ByVal vstrURL As String)
'************************************
'purpose:extracts email address from
'        URL indicated
'inputs:vstrSearchString-string to search on
'   ex: Javascript
'explanation:NONE
'returns:NONE
'************************************
'Dim intFreeFile As Integer
Dim lngPageNumber As Long
Dim blnMorePages As Boolean

Dim strNextUrl As String

'Dim strHtml As String
Dim rtbHtml As RichTextBox
Dim colUrl As Collection
Dim strLinkURL As Variant
Dim strUrl As Variant
Dim objNode As Node
Dim blnStayInDomain As Boolean

    '*******************************
    'disable page
    '*******************************
    frmWebBot.MousePointer = vbHourglass
    EnableForm frmWebBot, False
    tvLinks.Nodes.Clear
    
        

    mCrawlUrlForEmailAddresses vstrURL, "", 0, _
            txtSearchDepth - 1, 3, 0
            
 
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

Private Sub Form_Initialize()
    mblnReview = False
End Sub

Private Sub Form_Load()
    Center_Form Me
    mResetAllVisiting
    Me.Height = 6885
End Sub
Sub Center_Form(frmForm As Form)

       frmForm.Left = (Screen.Width - frmForm.Width) / 2
       frmForm.Top = (Screen.Height - frmForm.Height) / 2
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    'DoEvents
    'MsgBox "Internet control state = " & State

End Sub

Private Sub Inet2_StateChanged(ByVal State As Integer)
' Retrieve server response using the GetChunk
    ' method when State = 12. This example assumes the
    ' data is text.

    Select Case State
    ' ... Other cases not shown.

    Case icResponseReceived ' 12
        Dim vtData As Variant ' Data variable.
        Dim strData As String: strData = ""
        Dim bDone As Boolean: bDone = False

        ' Get first chunk.
        vtData = Inet1.GetChunk(1024, icString)
        DoEvents

Do While Not bDone

            strData = strData & vtData
            ' Get next chunk.
            vtData = Inet1.GetChunk(1024, icString)
            DoEvents

            If Len(vtData) = 0 Then
                bDone = True
            End If
        Loop

        txtHtml.Text = strData
        gboolFinishedReceived = True
    End Select
    

End Sub


Public Property Get Review() As Boolean
    Review = mblnReview
End Property

Public Property Let Review(ByVal vNewValue As Boolean)
    mblnReview = vNewValue
End Property

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuPurgeVisitedList_Click()
    mDeleteVisitedUrls
End Sub

Private Sub mnuReview_Click()
    'frmReviewAddresses.Show
End Sub

Private Sub mnuShowVisitedList_Click()
    'frmVisitedList.Show
End Sub

Private Sub tvLinks_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If tvLinks.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Button = 2 Then
        
        'right click
        'launch in browser
        'Shell "explorer.exe " & tvLinks.SelectedItem.Text
        frmWebBot.brwsWebBot.Navigate tvLinks.SelectedItem.Text
    Else
        'left click
        'tvLinks.SelectedItem.Expanded = Not (tvLinks.SelectedItem.Expanded)
    End If
End Sub
