VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReceipt 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receipt"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTFBox 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8493
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmReceipt.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
