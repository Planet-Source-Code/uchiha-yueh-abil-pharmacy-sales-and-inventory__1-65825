VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   1935
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   495
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3375
      Begin VB.TextBox txtPass 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   10
         PasswordChar    =   "="
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtUName 
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   795
         Width           =   690
      End
      Begin VB.Label lblUName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdOK_Click()

'check if user enter username..
If txtUName.Text = Empty Then
    MsgBox "Please enter your Username.", vbCritical, "Required"
    txtUName.SetFocus
    Exit Sub
End If

'check if user enter password..
If txtPass.Text = Empty Then
    MsgBox "Please enter your Password.", vbCritical, "Required"
    txtPass.SetFocus
    Exit Sub
End If

With Data1.Recordset
    
    'check if username exist..
    
    .MoveFirst
    .FindFirst "Username='" & txtUName.Text & "'"
    
    'if username does not exist..
    If .NoMatch = True Then
        MsgBox "Username does not exist.", vbCritical, "Not found"
        txtUName.SetFocus
        Exit Sub
    End If
    
    'check if password is the same in the records..
    If txtPass.Text <> .Fields("Password") Then
        'deny user..
        MsgBox "Access is denied.", vbCritical, Me.Caption
        
        txtUName.Text = Empty
        txtPass.Text = Empty
        txtUName.SetFocus
        Exit Sub
    
    End If
    
    'get username..
    User = .Fields("Username")
    
    'get usertype..
    UserType = .Fields("UserType")
    
    'get password..
    UPass = .Fields("Password")
    
    'display message..
    MsgBox "You are now login " & _
        UCase(.Fields("Username")), _
        vbInformation, Me.Caption

    Call Menus(UserType)

    Unload Me

End With

End Sub

Private Sub Form_Load()

With Data1

    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "User"
    .Refresh

End With

End Sub

Private Sub txtPass_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = "'" Then
    
    'When user enter a single quote ('),
    'no characters will be displayed..
    KeyAscii = 0

    'If keyascii is equal to 0,
    'no characters will be displayed
    'in the textbox..

ElseIf KeyAscii = vbKeyReturn Then
        
    Call cmdOK_Click

End If

End Sub

Private Sub txtUName_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtUName_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = "'" Then
    
    'When user enter a single quote ('),
    'no characters will be displayed..
    KeyAscii = 0

    'If keyascii is equal to 0,
    'no characters will be displayed
    'in the textbox..
    
ElseIf KeyAscii = vbKeyReturn Then

    'set focus to txtPass textbox..
    txtPass.SetFocus

End If

End Sub
