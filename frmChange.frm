VERSION 5.00
Begin VB.Form frmChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4275
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2250
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   810
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtConfirm 
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
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "="
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtNew 
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
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "="
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtOld 
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
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "="
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confir&m Password"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1275
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&New Password"
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
         Left            =   120
         TabIndex        =   2
         Top             =   795
         Width           =   1065
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old &Password"
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
         Left            =   120
         TabIndex        =   0
         Top             =   315
         Width           =   975
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdOK_Click()

'check if user enter old password..
If txtOld.Text = Empty Then
    MsgBox "Please enter your Old Password.", vbCritical, "Required"
    txtOld.SetFocus
    Exit Sub
End If

'check if user enter new password..
If txtNew.Text = Empty Then
    MsgBox "Please enter your New Password.", vbCritical, "Required"
    txtNew.SetFocus
    Exit Sub
End If

'check if user confirm password..
If txtConfirm.Text = Empty Then
    MsgBox "Please re-enter your new Password.", vbCritical, "Required"
    txtConfirm.SetFocus
    Exit Sub
End If

'check if old password is correct..
If txtOld.Text <> UPass Then
    MsgBox "Old Password password is incorrect.", vbCritical, "Message"
    txtOld.SetFocus
    Exit Sub
End If

'check if New and confirm password is the same..
If txtNew.Text <> txtConfirm.Text Then
    MsgBox "New and Confirm Password password is not the same.", vbCritical, "Message"
    txtNew.Text = Empty
    txtConfirm.Text = Empty
    txtNew.SetFocus
    Exit Sub
End If

With Data1.Recordset
    
    'edit password of user..
    .Edit
    
    .Fields("Password") = txtNew.Text
    
    'saving..
    .Update

End With

MsgBox "Password has been changed.", vbInformation, "Successful"

'get new password..
UPass = txtNew.Text

'call Click events of cmdClose..
Call cmdClose_Click

End Sub

Private Sub Form_Load()

'show toolbar..
mdiMain.Toolbar1.Visible = True

Me.Left = 0
Me.Top = 0

With Data1

    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "User"
    .Refresh
    
    With .Recordset
    
        .MoveFirst
        .FindFirst "[Username]='" & User & "'"
    
    End With

End With

End Sub

Private Sub txtConfirm_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = "'" Then
    
    'When user enter a single quote ('),
    'no characters will be displayed..
    KeyAscii = 0

    'If keyascii is equal to 0,
    'no characters will be displayed
    'in the textbox..
    
ElseIf KeyAscii = vbKeyReturn Then
    
    'call the Click Events of cmdOK..
    Call cmdOK_Click

End If

End Sub

Private Sub txtNew_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtNew_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = "'" Then
    
    'When user enter a single quote ('),
    'no characters will be displayed..
    KeyAscii = 0

    'If keyascii is equal to 0,
    'no characters will be displayed
    'in the textbox..
    
ElseIf KeyAscii = vbKeyReturn Then
    
    'set focus to txtConfirm textbox..
    txtConfirm.SetFocus

End If

End Sub

Private Sub txtOld_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = "'" Then
    
    'When user enter a single quote ('),
    'no characters will be displayed..
    KeyAscii = 0

    'If keyascii is equal to 0,
    'no characters will be displayed
    'in the textbox..
    
ElseIf KeyAscii = vbKeyReturn Then
    
    'set focus to txtNew textbox..
    txtNew.SetFocus

End If

End Sub
