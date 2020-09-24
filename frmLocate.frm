VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmLocate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Location"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLocate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5175
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
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   375
      Left            =   3082
      TabIndex        =   4
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   375
      Left            =   1597
      TabIndex        =   1
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   2587
      TabIndex        =   3
      Top             =   3000
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4935
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmLocate.frx":08CA
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frmLocate.frx":08DE
         TabIndex        =   13
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtLocate 
         DataField       =   "Location"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   0
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location of Product"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H8000000A&
      Caption         =   "&Add"
      Height          =   615
      Left            =   120
      Picture         =   "frmLocate.frx":1115
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3540
      Width           =   705
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000A&
      Caption         =   "&Edit"
      Height          =   615
      Left            =   825
      Picture         =   "frmLocate.frx":1207
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3540
      Width           =   825
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H8000000A&
      Caption         =   "&Delete"
      Height          =   615
      Left            =   1650
      Picture         =   "frmLocate.frx":12F9
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3540
      Width           =   825
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000A&
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2475
      Picture         =   "frmLocate.frx":13EB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3540
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000A&
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3300
      Picture         =   "frmLocate.frx":14DD
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3540
      Width           =   825
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000A&
      Caption         =   "Close"
      Height          =   615
      Left            =   4125
      Picture         =   "frmLocate.frx":15CF
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3540
      Width           =   855
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   2092
      TabIndex        =   2
      Top             =   3000
      Width           =   495
   End
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
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmLocate"
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

'set focus fo Location..
txtLocate.SetFocus

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
    
    'if there is no product..
    If .EOF = True And .RecordCount = 0 Then
        MsgBox "Product Location is empty.", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    'confirm deletion..
    ans = MsgBox("Delete '" & txtLocate.Text & "'?", _
                vbYesNo + vbQuestion, _
                "Confirm Deletion")
    
    If ans = vbNo Then Exit Sub
    
    .Delete

    MsgBox "Product Location has been deleted.", vbInformation, Me.Caption

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

'set focus fo Location..
txtLocate.SetFocus

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

'check Location..
If txtLocate.Text = Empty Then
    MsgBox "Enter Location of product.", vbExclamation, "Required"
    txtLocate.SetFocus
    Exit Sub
End If

'check if Location already exist..

If UserAction = "Add" Then
    
    With Data2.Recordset
    
        .MoveFirst
        .FindFirst "[Location]='" & txtLocate.Text & "'"
        
        'Location already exist..
        If .NoMatch = False Then
        
            MsgBox "Please enter a new Location.", _
                vbExclamation, "Location already exist."
                
            txtLocate.SetFocus
            SendKeys "{home}+{end}"
                
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
    
    'open database of Location..
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "Location"
    .Refresh
    
End With

With Data2
    
    'open database of Location..
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "Location"
    .Refresh
    
End With

DBGrid1.ReBind
DBGrid1.Refresh

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

txtLocate.Locked = Not txtLocate.Locked

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdClose.Enabled = False Then Cancel = 1

End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)

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
