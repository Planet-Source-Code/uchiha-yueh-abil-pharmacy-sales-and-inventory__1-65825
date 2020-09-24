VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Report"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7005
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Preview"
      Enabled         =   0   'False
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
      Left            =   5280
      Picture         =   "frmReport.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5280
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Result"
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6735
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmReport.frx":1994
         Height          =   3135
         Left            =   120
         OleObjectBlob   =   "frmReport.frx":19A8
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.PictureBox picDisplay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   120
         Picture         =   "frmReport.frx":2BE7
         ScaleHeight     =   3135
         ScaleWidth      =   6495
         TabIndex        =   18
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000A&
      Caption         =   "Close"
      Height          =   735
      Left            =   6120
      Picture         =   "frmReport.frx":4517D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2566
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "By Date"
      TabPicture(0)   =   "frmReport.frx":4526F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DTPicker1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCRep1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "By Month"
      TabPicture(1)   =   "frmReport.frx":4528B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdCRep2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboMonth"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtYear"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "By Range"
      TabPicture(2)   =   "frmReport.frx":452A7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "dtTo"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "dtFrom"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdCRep3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdCRep3 
         Caption         =   "&Create Report"
         Height          =   375
         Left            =   -70560
         TabIndex        =   16
         Top             =   645
         Width           =   2175
      End
      Begin VB.TextBox txtYear 
         Height          =   375
         Left            =   -73560
         MaxLength       =   4
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cboMonth 
         Height          =   345
         ItemData        =   "frmReport.frx":452C3
         Left            =   -73560
         List            =   "frmReport.frx":452EB
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   420
         Width           =   1815
      End
      Begin VB.CommandButton cmdCRep2 
         Caption         =   "&Create Report"
         Height          =   375
         Left            =   -70560
         TabIndex        =   8
         Top             =   645
         Width           =   2175
      End
      Begin VB.CommandButton cmdCRep1 
         Caption         =   "&Create Report"
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   645
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   645
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19267585
         CurrentDate     =   38418
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   375
         Left            =   -74280
         TabIndex        =   12
         Top             =   645
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19267585
         CurrentDate     =   38418
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   375
         Left            =   -72240
         TabIndex        =   14
         Top             =   645
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19267585
         CurrentDate     =   38418
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   225
         Left            =   -72720
         TabIndex        =   15
         Top             =   720
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   225
         Left            =   -74880
         TabIndex        =   13
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Year"
         Height          =   225
         Left            =   -74760
         TabIndex        =   10
         Top             =   915
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Month"
         Height          =   225
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a date"
         Height          =   225
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1065
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
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
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdCRep1_Click()

With Data1

    .RecordSource = _
        "SELECT * FROM Transaction WHERE Date=#" & _
        DTPicker1.Value & _
        "# ORDER BY [Product Code],[Product Name],Date"
    
    .Refresh
    
    With .Recordset
    
        If .EOF = True And .RecordCount = 0 Then
            
            MsgBox "No Record(s).", vbExclamation, Me.Caption
            cmdPrint.Enabled = False
            
            'hide grid..
            picDisplay.Visible = True
            DBGrid1.Visible = False
            
            Exit Sub
        
        Else
            
            picDisplay.Visible = False
            DBGrid1.Visible = True
        
        End If
    
    End With

End With

Frame1.Caption = "Result (By Date)"
cmdPrint.Enabled = True

End Sub

Private Sub cmdCRep2_Click()

'user did not select a month on the cboMonth..
If cboMonth.ListIndex = -1 Then

    MsgBox "Select a month.", vbExclamation, "Required"
    cboMonth.SetFocus
    Exit Sub

End If

'check year..
If txtYear.Text = Empty Then

    MsgBox "Enter year.", vbExclamation, "Required"
    txtYear.SetFocus
    Exit Sub

End If

With Data1

    .RecordSource = _
        "SELECT * FROM Transaction WHERE Month=" & _
        cboMonth.ListIndex + 1 & _
        " AND Year=" & Val(txtYear.Text) & _
        " ORDER BY Date,[Product Code],[Product Name]"
    
    .Refresh
    
    With .Recordset
    
        If .EOF = True And .RecordCount = 0 Then
            
            'display message..
            MsgBox "No Record(s).", vbExclamation, Me.Caption
            
            'disable print button..
            cmdPrint.Enabled = False
            
            'hide grid..
            picDisplay.Visible = True
            DBGrid1.Visible = False
            
            Exit Sub
        
        Else
        
            'show grid..
            picDisplay.Visible = False
            DBGrid1.Visible = True
        
        End If
    
    End With

End With

Frame1.Caption = "Result (By Month)"
cmdPrint.Enabled = True

End Sub

Private Sub cmdCRep3_Click()

'if FROM value is equal than TO value
If dtFrom.Value = dtTo.Value Then
    
    MsgBox "The date in FROM and TO is invalid.", vbExclamation, Me.Caption
    dtFrom.SetFocus
    Exit Sub
    
End If

With Data1

    .RecordSource = _
        "SELECT * FROM Transaction WHERE Date BETWEEN #" & dtFrom.Value & _
        "# AND #" & dtTo.Value & _
        "# ORDER BY Date,[Product Code],[Product Name]"
    
    .Refresh
    
    With .Recordset
    
        If .EOF = True And .RecordCount = 0 Then
            
            'display message..
            MsgBox "No Record(s).", vbExclamation, Me.Caption
            
            'disable print button..
            cmdPrint.Enabled = False
            
            'hide grid..
            picDisplay.Visible = True
            DBGrid1.Visible = False
            
            Exit Sub
        
        Else
        
            'show grid..
            picDisplay.Visible = False
            DBGrid1.Visible = True
        
        End If
    
    End With

End With

Frame1.Caption = "Result (By Range)"
cmdPrint.Enabled = True

End Sub

Private Sub cmdPrint_Click()

With Adodc1
    
    .RecordSource = Data1.RecordSource
    .Refresh
    
End With

Set drDaily.DataSource = Adodc1
ActiveReport = "Report"

'Call HideAll

End Sub


Private Sub Form_Load()

DTPicker1.Value = Date

dtFrom.Value = Date
dtTo.Value = Date

'show toolbar..
mdiMain.Toolbar1.Visible = True

Me.Left = 0
Me.Top = 0

With Adodc1

    .ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.3.51;" & _
        "Persist Security Info=False;" & _
        "Data Source=" & App.Path & "\database.mdb"

End With

With Data1
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "Transaction"
    .Refresh
End With

DBGrid1.Refresh
DBGrid1.ReBind

End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    
    Case 48 To 57      '0-9
    Case vbKeyBack
    Case Else
        KeyAscii = 0

End Select

End Sub
