VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   360
      Top             =   3000
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ABIL Pharmacy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2145
         TabIndex        =   6
         Top             =   720
         Width           =   2145
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   240
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   435
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrights 2005 © All Rights Reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4155
         TabIndex        =   3
         Top             =   3060
         Width           =   2820
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ABIL Pharmacy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning: Unauthorized reproduction or distribution of this product is strictly prohibited."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5580
         TabIndex        =   4
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Sales and Inventory System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   2160
         TabIndex        =   5
         Top             =   1080
         Width           =   4590
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()

Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'if user press the ESC..
If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

Me.Hide

mdiMain.Show
frmLogin.Show vbModal

End Sub

Private Sub Frame1_Click()

Unload Me

End Sub

Private Sub imgLogo_Click()

Unload Me

End Sub

Private Sub lblProductName_Click()

Unload Me

End Sub

Private Sub Timer1_Timer()

Unload Me

End Sub
