VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "NAT TRANS KeepAlive"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   345
      ItemData        =   "Form1.frx":000C
      Left            =   7200
      List            =   "Form1.frx":0025
      TabIndex        =   5
      Text            =   "1"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2400
      TabIndex        =   3
      Top             =   2610
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   2610
      Width           =   855
   End
   Begin InetCtlsObjects.Inet in1 
      Left            =   3480
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser w 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      ExtentX         =   16325
      ExtentY         =   2778
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer log 
      Interval        =   1000
      Left            =   1200
      Top             =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Times As Integer, MyCounter As Integer, WebSite(20) As String, MaxSites As Integer

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

MaxSites = 10 ' # of web sites through which to cycle, list them next
WebSite(1) = "http://www.yahoo.com"
WebSite(2) = "http://www.altavista.com"
WebSite(3) = "http://www.lycos.com"
WebSite(4) = "http://www.webcrawler.com"
WebSite(5) = "http://www.askjeeves.com"
WebSite(6) = "http://www.winzip.com"
WebSite(7) = "http://www.ipswitch.com"
WebSite(8) = "http://www.cuteftp.com"
WebSite(9) = "http://www.tucows.com"
WebSite(10) = "http://maps.yahoo.com"

Call Form_Resize ' set screen location for objects

Times = 0
MyCounter = 0
log.Interval = 1000
Text1.Alignment = 2
Combo1.Text = 120 ' default to 2 hours between page changes

w.Navigate WebSite(MaxSites) ' load last site without waiting
in1.OpenURL WebSite(MaxSites) ' load last site without waiting

End Sub

Private Sub Form_Resize()

On Error GoTo BailOut

If Form1.WindowState <> vbMinimized Then

    ' set browser dimensions
    w.Left = 50
    w.Top = 50
    w.Width = Me.Width - 200
    w.Height = Me.Height - 1000

    ' set vertical location of text and buttons
    Text1.Top = Me.Height - 850
    Text2.Top = Text1.Top
    Label1.Top = Text1.Top
    Combo1.Top = Text1.Top
    Command2.Top = Text1.Top

    ' set horizontal location of text and buttons
    Label1.Left = 50
    Text1.Left = Label1.Left + Label1.Width + 50
    Text2.Left = Text1.Left + Text1.Width + 50
    Combo1.Left = Text2.Left + Text2.Width + 50
    Command2.Left = Combo1.Left + Combo1.Width + 50
    
    ' set height of text and buttons
    Text1.Height = Label1.Height
    Text2.Height = Label1.Height
    Combo1.Height = Label1.Height
    Command2.Height = Label1.Height

End If

BailOut: ' exit if screen dimensions fail

End Sub

Private Sub log_Timer()

Label1.Caption = Time$
MyCounter = MyCounter + 1
Form1.Caption = "NTKA - " & CStr(MyCounter) & "/" & CStr(Times)

If MyCounter = CInt(Combo1.Text) * 60 Then ' # of minutes between site loads

    Times = Times + 1
    w.Navigate WebSite(Times)
    in1.OpenURL WebSite(Times)
    If Times = MaxSites Then Times = 0
    MyCounter = 0
    
End If

Text1.Text = CStr(MyCounter) & "/" & CStr(Times)
Text2.Text = " " & w.LocationURL
Form1.Refresh
Text1.Refresh
Text2.Refresh

End Sub

