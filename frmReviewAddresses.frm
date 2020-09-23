VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmReviewAddresses 
   Caption         =   "Review Email Addresses"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   9015
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin MSAdodcLib.Adodc adoEmailAddresses 
         Height          =   2895
         Left            =   120
         Top             =   240
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   5106
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   1
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * from Extracted_Email_address "
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dgrdEmailAddress 
         Bindings        =   "frmReviewAddresses.frx":0000
         Height          =   2895
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5106
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "Preview:"
         Height          =   5775
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   9375
         Begin SHDocVwCtl.WebBrowser webbrwsEmailAddress 
            Height          =   5415
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   9135
            ExtentX         =   16113
            ExtentY         =   9551
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
      End
      Begin VB.CommandButton cmClose 
         Caption         =   "&Close"
         Height          =   495
         Left            =   9720
         TabIndex        =   1
         Top             =   6240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReviewAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrURL As String
Private mblnDisableChangeEvent As Boolean

Private Sub cmClose_Click()
    Unload Me
End Sub

Private Sub cmdSQL_Click()
Dim objConnection As ADODB.Connection

'    If (MsgBox("Are you sure you run this SQL statement?", vbYesNo + vbDefaultButton1, _
'        "Purge?") = vbYes) Then
'
'        'connect to database
'        ConnectToDatabase objConnection
'
'        'open recordset
'        objConnection.Execute txtSQL
'
'        DisconnectFromDatabase objConnection
'
'        'refresh data
'        'dgrdEmailAddress.Refresh
'        mRequeryDataGrid
'
'    End If

End Sub

Private Sub cmNext_Click()
    dgrdEmailAddress.Row = dgrdEmailAddress.Row + 1
   ' dgrdEmailAddress.SelStartCol = 1
    'dgrdEmailAddress.SelEndCol = 5
End Sub

Private Sub cmShowInBrowser_Click()
    Shell "explorer.exe " & mstrURL
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dgrdEmailAddress_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim intCurrentRow As Integer
Dim intLastRow As Integer

    If (mblnDisableChangeEvent = False) Then
        'update last cell as reviewed
        On Error GoTo lblSkip
        intLastRow = LastRow
        On Error GoTo 0
'
'        intCurrentRow = dgrdEmailAddress.Row
'
'        mblnDisableChangeEvent = True
'        dgrdEmailAddress.Row = intLastRow
'        mblnDisableChangeEvent = False
        
        'dgrdEmailAddress.Columns.Item(5).Value = 1
        
'        mblnDisableChangeEvent = True
'        dgrdEmailAddress.Row = intCurrentRow
'        mblnDisableChangeEvent = False
        
    
        'user moved to new cell
        'show in browser
lblSkip:
        mstrURL = dgrdEmailAddress.Columns(2).Value
        webbrwsEmailAddress.Navigate mstrURL

    End If
    
End Sub

Private Sub optReview_Click(Index As Integer)
    
    mRequeryDataGrid
    
End Sub

Private Sub optWorld_Click(Index As Integer)

    mRequeryDataGrid
    
End Sub


Sub mRequeryDataGrid()
'***************************************
'purpose:requeries data grid
'inputs:NONE
'returns:NONE
'explanation:NONE
'***************************************
'    Me.MousePointer = vbHourglass
'   ' EnableForm Me, False
'   dgrdEmailAddress.Enabled = False
'    adoEmailAddresses.RecordSource = "select * from Extracted_email_address " & vbCrLf & _
'        " where reviewed = " & mintSelectedControlId(optReview) & vbCrLf & _
'        " ORDER BY ExtractedEmailAddressId"
'    DoEvents
'    adoEmailAddresses.Refresh
'    DoEvents
'    'EnableForm Me, True
'    dgrdEmailAddress.Enabled = True
'    Me.MousePointer = vbNormal
    
End Sub

Private Sub Form_Load()
    
    adoEmailAddresses.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/" & "WebAgent.mdb;Persist Security Info=False"

End Sub
