VERSION 5.00
Begin VB.Form frmUser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Account"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5115
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   2085
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000A&
      Caption         =   "Close"
      Height          =   615
      Left            =   4125
      Picture         =   "frmUser.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2580
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000A&
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3300
      Picture         =   "frmUser.frx":0DBC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2580
      Width           =   825
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000A&
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2475
      Picture         =   "frmUser.frx":0EAE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2580
      Width           =   825
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H8000000A&
      Caption         =   "&Delete"
      Height          =   615
      Left            =   1650
      Picture         =   "frmUser.frx":0FA0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2580
      Width           =   825
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000A&
      Caption         =   "&Edit"
      Height          =   615
      Left            =   825
      Picture         =   "frmUser.frx":1092
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2580
      Width           =   825
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H8000000A&
      Caption         =   "&Add"
      Height          =   615
      Left            =   120
      Picture         =   "frmUser.frx":1184
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2580
      Width           =   705
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   2580
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   375
      Left            =   1590
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   375
      Left            =   3075
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtUType 
         DataField       =   "UserType"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         DataField       =   "Name"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   3
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtUName 
         DataField       =   "Username"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         Top             =   285
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User &Type"
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Top             =   1275
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Name"
         Height          =   225
         Left            =   240
         TabIndex        =   2
         Top             =   795
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Username"
         Height          =   225
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserAction As String

Private Sub cmdAdd_Click()

'disable buttons..
Call Controlls

'AddNew..
Data1.Recordset.AddNew

'set Usertype..
txtUType.Text = "USER"

'set focus fo txtUname..
txtUName.SetFocus

End Sub

Private Sub cmdCancel_Click()

'Cancel..
Data1.Recordset.CancelUpdate

'enable buttons..
Call Controlls

End Sub

Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdDelete_Click()

With Data1.Recordset
    
    'if user is "admin"..
    If .Fields("UserType") = "Administrator" Then
        MsgBox "User Administrator cannot be deleted.", _
            vbExclamation, "Access is denied"
        Exit Sub
    End If
    
    'confirm deletion..
    ans = MsgBox("Delete '" & txtUName.Text & "'?", _
                vbYesNo + vbQuestion, _
                "Confirm Deletion")
    
    If ans = vbNo Then Exit Sub
    
    .Delete

    MsgBox "User has been deleted.", vbInformation, Me.Caption

End With

Data1.Refresh

End Sub

Private Sub cmdEdit_Click()

'Edit..
With Data1.Recordset

    If .RecordCount = 0 And .EOF = True Then
        MsgBox "No record(s)", vbCritical, "Edit record failed"
        Exit Sub
    End If
    
    .Edit

End With

'disable buttons..
Call Controlls

'set focus fo txtUName..
txtUName.SetFocus

End Sub

Private Sub cmdFirst_Click()

On Error Resume Next

With Data1.Recordset
    
    .MoveFirst

End With

End Sub

Private Sub cmdLast_Click()

On Error Resume Next

With Data1.Recordset
    
    .MoveLast

End With

End Sub

Private Sub cmdNext_Click()

On Error Resume Next

With Data1.Recordset
    
    .MoveNext
    If .EOF = True Then .MoveLast

End With

End Sub

Private Sub cmdPrev_Click()

On Error Resume Next

With Data1.Recordset
    
    .MovePrevious
    
    If .BOF = True Then .MoveFirst

End With

End Sub

Private Sub cmdSave_Click()

'check Username..
If txtUName.Text = Empty Then
    MsgBox "Enter Username.", vbExclamation, "Required"
    txtUName.SetFocus
    Exit Sub
End If

'check if Username already exist..

If UserAction = "Add" Then
    
    With Data2.Recordset
    
        .MoveFirst
        .FindFirst "[Username]='" & txtUName.Text & "'"
        
        'Username already exist..
        If .NoMatch = False Then
        
            MsgBox "Please enter a new Username.", _
                vbExclamation, "Username already exist."
                
            txtUName.SetFocus
            Exit Sub
            
        End If
    
    End With

End If

'check Name..
If txtName.Text = Empty Then
    MsgBox "Enter Name.", vbExclamation, "Required"
    txtName.SetFocus
    Exit Sub
End If

'check if Name already exist..

If UserAction = "Add" Then
    
    With Data2.Recordset
    
        .MoveFirst
        .FindFirst "[Name]='" & txtName.Text & "'"
        
        'Name already exist..
        If .NoMatch = False Then
        
            MsgBox "Please enter a new Name.", _
                vbExclamation, "Name already exist."
                
            txtName.SetFocus
            Exit Sub
            
        End If
    
    End With

End If

'saving..
Data1.Recordset.Update

'enable controls..
Call Controlls

End Sub

Private Sub Form_Load()

'show toolbar..
mdiMain.Toolbar1.Visible = True

Me.Left = 0
Me.Top = 0
    
With Data1
    
    'open database of User..
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "User"
    .Refresh
    
End With

With Data2
    
    'open database of User..
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "User"
    .Refresh
    
End With

End Sub

Sub Controlls()

cmdFirst.Enabled = Not cmdFirst.Enabled
cmdPrev.Enabled = Not cmdPrev.Enabled
cmdNext.Enabled = Not cmdNext.Enabled
cmdLast.Enabled = Not cmdLast.Enabled

cmdAdd.Enabled = Not cmdAdd.Enabled
cmdEdit.Enabled = Not cmdEdit.Enabled
cmdDelete.Enabled = Not cmdDelete.Enabled
cmdSave.Enabled = Not cmdSave.Enabled
cmdCancel.Enabled = Not cmdCancel.Enabled
cmdClose.Enabled = Not cmdClose.Enabled

txtUName.Locked = Not txtUName.Locked
txtName.Locked = Not txtName.Locked

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdClose.Enabled = False Then Cancel = 1

End Sub

Private Sub txtName_GotFocus()

SendKeys "{home}+{end}"

End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 97 To 122
        
        'a-z, it changes from lowercase to upper case..
        KeyAscii = KeyAscii - 32
    
    Case 39
    
        'When user enter a single quote ('),
        'no characters will be displayed..
        KeyAscii = 0

        'If keyascii is equal to 0,
        'no characters will be displayed
        'in the textbox..

End Select

End Sub

Private Sub txtUName_GotFocus()

SendKeys "{home}+{end}"

End Sub

Private Sub txtUName_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    
    Case 39
    
        'When user enter a single quote ('),
        'no characters will be displayed..
        KeyAscii = 0

        'If keyascii is equal to 0,
        'no characters will be displayed
        'in the textbox..

End Select

End Sub

Private Sub txtUType_GotFocus()

SendKeys "{home}+{end}"

End Sub
