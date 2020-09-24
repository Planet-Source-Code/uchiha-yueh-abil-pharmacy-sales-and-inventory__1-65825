VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmExpire 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expired Products"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExpire.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8430
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Picture         =   "frmExpire.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000A&
      Caption         =   "Close"
      Height          =   735
      Left            =   7575
      Picture         =   "frmExpire.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Align           =   1  'Align Top
      Bindings        =   "frmExpire.frx":1286
      Height          =   4215
      Left            =   0
      OleObjectBlob   =   "frmExpire.frx":129A
      TabIndex        =   0
      Top             =   0
      Width           =   8430
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3480
      Top             =   2520
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmExpire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdPrint_Click()

With Adodc1

    .ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.3.51;" & _
        "Persist Security Info=False;" & _
        "Data Source=" & App.Path & "\database.mdb"

    .RecordSource = _
        "SELECT * FROM Product WHERE Product.Month < " & _
        Month(Now) & " AND Product.Year < " & Year(Now) & _
        " ORDER BY Product.Month, Product.Year"

    .Refresh

End With

Set drExpire.DataSource = Adodc1
ActiveReport = "Expire"
'Call HideAll

End Sub

Private Sub Form_Load()

'show toolbar..
mdiMain.Toolbar1.Visible = True

Me.Left = 0
Me.Top = 0

With Data1

    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = _
        "SELECT * FROM Product WHERE Month < " & _
        Month(Now) & " AND Year < " & Year(Now) & _
        " ORDER BY Month, Year"
    
    .Refresh
    
End With

End Sub
