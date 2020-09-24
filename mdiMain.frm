VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "ABIL Pharmacy -  Sales and Inventory System"
   ClientHeight    =   6660
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10830
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   1429
      ButtonWidth     =   2461
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "POS"
            Key             =   "POS"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Products"
            Key             =   "Products"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Report"
            Key             =   "Report"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Change Password"
            Key             =   "Password"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0CCA
            Key             =   "POS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":19A4
            Key             =   "Products"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":29F6
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":36D0
            Key             =   "Password"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6285
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13467
            Text            =   "User"
            TextSave        =   "User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/1/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "4:26 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu SP00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Report"
      End
   End
   Begin VB.Menu mnuInvent 
      Caption         =   "Inventory"
      Begin VB.Menu mnuLog 
         Caption         =   "Login"
         Shortcut        =   {F12}
      End
      Begin VB.Menu SP0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPOS 
         Caption         =   "Point of Sales"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu SP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProd 
         Caption         =   "Products"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExpired 
         Caption         =   "Expired Products"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLocate 
         Caption         =   "Product Location"
         Enabled         =   0   'False
      End
      Begin VB.Menu SP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuInvRept 
         Caption         =   "Sales Report"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuSecure 
      Caption         =   "Security"
      Begin VB.Menu mnuUser 
         Caption         =   "User Account"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuChange 
         Caption         =   "Change Password"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuSIS 
         Caption         =   "About ABIL Pharmacy -  Sales and Inventory System"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuChange_Click()

frmChange.Show

End Sub

Private Sub mnuClose_Click()

If ActiveReport = "Report" Then Unload drDaily
If ActiveReport = "Expire" Then Unload drExpire

End Sub

Private Sub mnuExit_Click()

End

End Sub

Private Sub mnuExpired_Click()

frmExpire.Show

End Sub

Private Sub mnuInvRept_Click()

frmReport.Show

End Sub

Private Sub mnuLocate_Click()

frmLocate.Show

End Sub

Private Sub mnuLog_Click()

'check mnuLog caption..
If Me.mnuLog.Caption = "Login" Then
    
    frmLogin.Show vbModal

Else

    If Screen.ActiveForm Is frmChange Then _
        Unload frmChange

    If Screen.ActiveForm Is frmExpire Then _
        Unload frmExpire

    If Screen.ActiveForm Is frmLocate Then _
        Unload frmLocate

    If Screen.ActiveForm Is frmPOS Then _
        Unload frmPOS

    If Screen.ActiveForm Is frmProduct Then _
        Unload frmProduct

    If Screen.ActiveForm Is frmUser Then _
        Unload frmUser

    'toggle menu buttons..
    Call Menus(Empty)

    'show UserLogin form..
    frmLogin.Show vbModal
    
End If

End Sub

Private Sub mnuPOS_Click()

frmPOS.Show

End Sub

Private Sub mnuPrint_Click()

If mdiMain.ActiveForm Is drDaily Then
    drDaily.PrintReport True
End If

End Sub

Private Sub mnuProd_Click()

frmProduct.Show

End Sub

Private Sub mnuSIS_Click()

msg = "ABIL Pharmacy -  Sales and Inventory System"
msg = msg & vbCrLf
msg = msg & vbCrLf
msg = msg & "Copyrights 2005 © All Rights Reserved"

MsgBox msg, vbInformation, "About ABIL Pharmacy -  Sales and Inventory System"

End Sub

Private Sub mnuUser_Click()

frmUser.Show

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case Is = "POS"
        Call mnuPOS_Click

    Case Is = "Products"
        Call mnuProd_Click

    Case Is = "Report"
        Call mnuInvRept_Click

    Case Is = "Password"
        Call mnuChange_Click

End Select

End Sub
