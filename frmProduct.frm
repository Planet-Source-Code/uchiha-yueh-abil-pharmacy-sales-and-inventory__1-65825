VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProduct 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Products"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProduct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6630
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   375
      Left            =   3810
      TabIndex        =   29
      Top             =   4665
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   6375
      Begin MSMask.MaskEdBox meExpire 
         DataField       =   "Expiration Date"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   4080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         Mask            =   "##/##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtReOrder 
         DataField       =   "Reorder Level"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   7
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtOnHand 
         DataField       =   "OnHand"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtBName 
         DataField       =   "Brand Name"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   2
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox txtGName 
         DataField       =   "Generic Name"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   1
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtPCode 
         DataField       =   "Product Code"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboUnit 
         Height          =   345
         ItemData        =   "frmProduct.frx":1042
         Left            =   2160
         List            =   "frmProduct.frx":1058
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtUnit 
         DataField       =   "Unit"
         DataSource      =   "Data1"
         Height          =   345
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cboLocation 
         Height          =   345
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3600
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtLocation 
         DataField       =   "Location"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3600
         Width           =   3375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   225
         Left            =   1755
         TabIndex        =   30
         Top             =   2235
         Width           =   330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Date (MM/YR)"
         Height          =   225
         Left            =   165
         TabIndex        =   28
         Top             =   4155
         Width           =   1965
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   225
         Left            =   1395
         TabIndex        =   27
         Top             =   3675
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Order Level"
         Height          =   225
         Left            =   840
         TabIndex        =   26
         Top             =   3195
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OnHand"
         Height          =   225
         Left            =   1395
         TabIndex        =   25
         Top             =   2715
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
         Height          =   225
         Left            =   1035
         TabIndex        =   24
         Top             =   1275
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generic Name"
         Height          =   225
         Left            =   885
         TabIndex        =   23
         Top             =   795
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   225
         Left            =   960
         TabIndex        =   22
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   225
         Left            =   1665
         TabIndex        =   21
         Top             =   1755
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   375
      Left            =   2325
      TabIndex        =   11
      Top             =   4665
      Width           =   495
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   2820
      TabIndex        =   12
      Top             =   4665
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   3315
      TabIndex        =   13
      Top             =   4665
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000A&
      Caption         =   "Close"
      Height          =   615
      Left            =   5400
      Picture         =   "frmProduct.frx":108B
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5205
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000A&
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4320
      Picture         =   "frmProduct.frx":117D
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5205
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000A&
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      Picture         =   "frmProduct.frx":126F
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5205
      Width           =   1065
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H8000000A&
      Caption         =   "&Delete"
      Height          =   615
      Left            =   2160
      Picture         =   "frmProduct.frx":1361
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5205
      Width           =   1065
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000A&
      Caption         =   "&Edit"
      Height          =   615
      Left            =   1080
      Picture         =   "frmProduct.frx":1453
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5205
      Width           =   1065
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H8000000A&
      Caption         =   "&Add"
      Height          =   615
      Left            =   120
      Picture         =   "frmProduct.frx":1545
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5205
      Width           =   945
   End
End
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboLocation_Click()

txtLocation.Text = cboLocation.Text

End Sub

Private Sub cboUnit_Click()

txtUnit.Text = cboUnit.Text

End Sub

Private Sub cmdAdd_Click()

'disable buttons..
Call Controlls

'hide Location textbox..
txtLocation.Visible = False

'show Location combo box..
cboLocation.Visible = True

'AddNew..
Data1.Recordset.AddNew

'set focus fo Product Code..
txtPCode.SetFocus

End Sub

Private Sub cmdCancel_Click()

'Cancel..
Data1.Recordset.CancelUpdate

'enable buttons..
Call Controlls

'show Location textbox..
txtLocation.Visible = True

'hide Location combo box..
cboLocation.Visible = False

End Sub

Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdDelete_Click()

With Data1.Recordset
    
    'if there is no product..
    If .EOF = True And .RecordCount = 0 Then
        MsgBox "Product is empty.", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    'confirm deletion..
    ans = MsgBox("Delete " & txtGName.Text & "?", _
                vbYesNo + vbQuestion, _
                "Confirm Deletion")
    
    If ans = vbNo Then Exit Sub
    
    .Delete

    MsgBox "Product has been deleted.", vbInformation, Me.Caption

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

'hide Location textbox..
txtLocation.Visible = False

'show Location combo box..
cboLocation.Visible = True

'search for Location Name..
For i = 0 To cboLocation.ListCount - 1
    If txtLocation.Text = cboLocation.List(i) Then
        cboLocation.ListIndex = i
        Exit For
    End If
Next

'search for unit..
For i = 0 To cboUnit.ListCount - 1
    If txtUnit.Text = cboUnit.List(i) Then
        cboUnit.ListIndex = i
        Exit For
    End If
Next

'set focus fo Product Code..
txtPCode.SetFocus

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

'check Product Code..
If txtPCode.Text = Empty Then
    MsgBox "Enter Product Name.", vbExclamation, "Required"
    txtPCode.SetFocus
    Exit Sub
End If

'check Generic Name..
If txtGName.Text = Empty Then
    MsgBox "Enter Generic Name.", vbExclamation, "Required"
    txtGName.SetFocus
    Exit Sub
End If

'check Brand Name..
If txtBName.Text = Empty Then
    MsgBox "Enter Brand Name.", vbExclamation, "Required"
    txtBName.SetFocus
    Exit Sub
End If

'check Price..
If txtPrice.Text = Empty Then
    MsgBox "Enter Price.", vbExclamation, "Required"
    txtPrice.SetFocus
    Exit Sub
End If

'check Unit..
If txtUnit.Text = Empty Then
    MsgBox "Select a unit.", vbExclamation, "Required"
    cboUnit.SetFocus
    Exit Sub
End If

'check OnHand..
If txtOnHand.Text = Empty Then
    MsgBox "Enter OnHand.", vbExclamation, "Required"
    txtOnHand.SetFocus
    Exit Sub
End If

'check ReOrder..
If txtReOrder.Text = Empty Then
    MsgBox "Enter Re-Order Level.", vbExclamation, "Required"
    txtReOrder.SetFocus
    Exit Sub
End If

'check Location..
If txtLocation.Text = Empty Then
    MsgBox "Select a  Location.", vbExclamation, "Required"
    cboLocation.SetFocus
    Exit Sub
End If

'check Expiration Date..
If meExpire.Text = "__/__" Then
    MsgBox "Enter Expiration Date.", vbExclamation, "Required"
    meExpire.SetFocus
    Exit Sub
End If

'check content of Expiration Date..
For i = 1 To Len(meExpire.Text)
    If Mid(meExpire.Text, i, 1) = "_" Then
        MsgBox "Complete Expiration Date.", vbExclamation, "Required"
        meExpire.SetFocus
        Exit Sub
    End If
Next

'check if the date in expiration date is correct..
Mth = Left(meExpire.Text, 2)
Yr = Right(meExpire.Text, 2)

If Mth > 12 Then
    MsgBox "Enter the correct date of Expiration Date.", _
        vbExclamation, "Error"
    meExpire.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

'saving..
Data1.Recordset.Update

'enable controls..
Call Controlls

'show Location textbox..
txtLocation.Visible = True

'hide Location combo box..
cboLocation.Visible = False

End Sub

Private Sub Form_Load()

'show toolbar..
mdiMain.Toolbar1.Visible = True

Me.Left = 0
Me.Top = 0

With Data2
    
    'open database of Location..
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "Location"
    .Refresh
    
    With .Recordset
    
        .MoveFirst
        
        Do While .EOF = False
        
            'add Location Name to cboLocation..
            cboLocation.AddItem ![Location]
            .MoveNext
            
        Loop
    
    End With
    
End With
    
With Data1
    
    'open database of Products..
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = _
        "SELECT * FROM Product ORDER BY [Product Code]," & _
        "[Generic Name],[Brand Name]"
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

txtPCode.Locked = Not txtPCode.Locked
txtGName.Locked = Not txtGName.Locked
txtBName.Locked = Not txtBName.Locked
txtPrice.Locked = Not txtPrice.Locked
txtOnHand.Locked = Not txtOnHand.Locked
txtReOrder.Locked = Not txtReOrder.Locked

txtUnit.Visible = Not txtUnit.Visible
cboUnit.Visible = Not cboUnit.Visible

txtLocation.Visible = Not txtLocation.Visible
cboLocation.Visible = Not cboLocation.Visible

meExpire.Enabled = Not meExpire.Enabled

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdClose.Enabled = False Then Cancel = 1

End Sub

Private Sub txtBName_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 32             'space
    Case 45, 47         'slash and hyphen
    Case 65 To 90       'A to Z..
    Case 97 To 122
    Case 48 To 57       '0-9
    Case vbKeyBack
    Case Else
        KeyAscii = 0

End Select

End Sub

Private Sub txtGName_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 32             'space
    Case 45, 47         'slash and hyphen
    Case 65 To 90       'A to Z..
    Case 97 To 122
    Case 48 To 57       '0-9
    Case vbKeyBack
    Case Else
        KeyAscii = 0

End Select

End Sub

Private Sub txtOnHand_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    
    Case 48 To 57       '0-9
    Case vbKeyBack
    Case Else
        KeyAscii = 0

End Select

End Sub

Private Sub txtPCode_Change()

txtPCode.Text = UCase(txtPCode.Text)

End Sub

Private Sub txtPCode_KeyPress(KeyAscii As Integer)

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

Private Sub txtPrice_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    
    Case 46, 48 To 57      '0-9
    Case vbKeyBack
    Case Else
        KeyAscii = 0

End Select

End Sub

Private Sub txtReOrder_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    
    Case 48 To 57       '0-9
    Case vbKeyBack
    Case Else
        KeyAscii = 0

End Select

End Sub
