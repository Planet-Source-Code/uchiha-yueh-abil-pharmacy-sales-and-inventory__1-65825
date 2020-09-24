VERSION 5.00
Begin VB.Form frmSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier"
   ClientHeight    =   3360
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
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
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
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   375
      Left            =   3082
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   375
      Left            =   1597
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   2587
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   4935
      Begin VB.TextBox txtSDesc 
         DataField       =   "Description"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtSName 
         DataField       =   "Supplier Name"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   1
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtSCode 
         DataField       =   "Supplier Code"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   1395
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         Height          =   225
         Left            =   240
         TabIndex        =   15
         Top             =   915
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
         Height          =   225
         Left            =   240
         TabIndex        =   14
         Top             =   435
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H8000000A&
      Caption         =   "&Add"
      Height          =   615
      Left            =   120
      Picture         =   "frmSupplier.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2580
      Width           =   705
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000A&
      Caption         =   "&Edit"
      Height          =   615
      Left            =   825
      Picture         =   "frmSupplier.frx":09BC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2580
      Width           =   825
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H8000000A&
      Caption         =   "&Delete"
      Height          =   615
      Left            =   1650
      Picture         =   "frmSupplier.frx":0AAE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2580
      Width           =   825
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000A&
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2475
      Picture         =   "frmSupplier.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2580
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000A&
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3300
      Picture         =   "frmSupplier.frx":0C92
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2580
      Width           =   825
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000A&
      Caption         =   "Close"
      Height          =   615
      Left            =   4125
      Picture         =   "frmSupplier.frx":0D84
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2580
      Width           =   855
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   2092
      TabIndex        =   4
      Top             =   2040
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
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmSupplier"
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

'set focus fo Supplier Code..
txtSCode.SetFocus

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
        MsgBox "Supplier is empty.", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    'confirm deletion..
    ans = MsgBox("Delete " & txtSCode.Text & "?", _
                vbYesNo + vbQuestion, _
                "Confirm Deletion")
    
    If ans = vbNo Then Exit Sub
    
    .Delete

    MsgBox "Supplier has been deleted.", vbInformation, Me.Caption

End With

Data1.Refresh

End Sub

Private Sub cmdEdit_Click()

'disable buttons..
Call Controlls

'Edit..
Data1.Recordset.Edit

'set focus fo Supplier Code..
txtSCode.SetFocus

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

'check Supplier Code..
If txtSCode.Text = Empty Then
    MsgBox "Enter Product Name.", vbExclamation, "Required"
    txtSCode.SetFocus
    Exit Sub
End If

'check if supplier code already exist..

If UserAction = "Add" Then
    
    With Data2.Recordset
    
        .MoveFirst
        .FindFirst "[Supplier Code]='" & txtSCode.Text & "'"
        
        'supplier code already exist..
        If .NoMatch = False Then
        
            MsgBox "Please enter a new Supplier Code.", _
                vbExclamation, "Supplier Code already exist."
                
            txtSCode.SetFocus
            SendKeys "{home}+{end}"
                
            Exit Sub
            
        End If
    
    End With

End If

'00------------------------------------------------00

'check Supplier Name..
If txtSName.Text = Empty Then
    MsgBox "Enter Supplier Name.", vbExclamation, "Required"
    txtSName.SetFocus
    Exit Sub
End If

'check if supplier name already exist..

If UserAction = "Add" Then
    
    With Data2.Recordset
    
        .MoveFirst
        .FindFirst "[Supplier Name]='" & txtSName.Text & "'"
        
        'supplier name already exist..
        If .NoMatch = False Then
        
            MsgBox "Please enter a new Supplier Name.", _
                vbExclamation, "Supplier Name already exist."
                
            txtSName.SetFocus
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
    
With Data1
    
    'open database of Products..
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "Supplier"
    .Refresh
    
End With

With Data2
    
    'open database of Suppplier..
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "Supplier"
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

txtSCode.Locked = Not txtSCode.Locked
txtSName.Locked = Not txtSName.Locked
txtSDesc.Locked = Not txtSDesc.Locked

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdClose.Enabled = False Then Cancel = 1

End Sub

Private Sub txtSCode_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 65 To 90       'A to Z..
    Case 97 To 122
        KeyAscii = KeyAscii - 32
    Case 48 To 57       '0-9
    Case vbKeyBack
    Case Else
        KeyAscii = 0

End Select

End Sub


Private Sub txtSName_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = "'" Or _
    Chr(KeyAscii) = "*" Then KeyAscii = 0

End Sub
