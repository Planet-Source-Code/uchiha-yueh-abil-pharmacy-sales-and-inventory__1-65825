VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPOS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Point of Sales"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11520
   Begin VB.PictureBox picAmount 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   2738
      Picture         =   "frmPOS.frx":08CA
      ScaleHeight     =   1500
      ScaleWidth      =   3360
      TabIndex        =   23
      Top             =   1852
      Visible         =   0   'False
      Width           =   3390
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   6015
         TabIndex        =   25
         Top             =   0
         Width           =   6015
         Begin VB.Label lblAmountTender 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   1380
         End
      End
      Begin VB.Frame fraAmount 
         Height          =   2415
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   5775
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   27
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            TabIndex        =   28
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtTender 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   120
            MaxLength       =   10
            TabIndex        =   26
            Top             =   240
            Width           =   5535
         End
      End
   End
   Begin VB.PictureBox picForm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   2093
      Picture         =   "frmPOS.frx":3EE1C
      ScaleHeight     =   4065
      ScaleWidth      =   3705
      TabIndex        =   21
      Top             =   1125
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Frame fraForm 
         Height          =   4935
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   7095
         Begin VB.Data dbPrice 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   2520
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Data dbProduct 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   1080
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdClose 
            Cancel          =   -1  'True
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5520
            TabIndex        =   44
            Top             =   4320
            Width           =   1455
         End
         Begin VB.CommandButton cmdInsert 
            Caption         =   "Insert"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   43
            Top             =   4320
            Width           =   1455
         End
         Begin MSDBGrid.DBGrid dbProd 
            Bindings        =   "frmPOS.frx":B6A26
            Height          =   2640
            Left            =   120
            OleObjectBlob   =   "frmPOS.frx":B6A3C
            TabIndex        =   42
            Top             =   1575
            Width           =   6855
         End
         Begin VB.TextBox txtSearch 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   1200
            Width           =   6855
         End
         Begin VB.ComboBox cboSearch 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmPOS.frx":B73FE
            Left            =   1200
            List            =   "frmPOS.frx":B740B
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Product Code / Brand Name / Generic Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   4125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search by"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   39
            Top             =   420
            Width           =   810
         End
      End
   End
   Begin VB.PictureBox picLocate 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   3233
      ScaleHeight     =   2385
      ScaleWidth      =   5025
      TabIndex        =   50
      Top             =   2250
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   120
         TabIndex        =   51
         Top             =   0
         Width           =   4815
         Begin VB.ComboBox cboProduct 
            Height          =   360
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   480
            Width           =   4575
         End
         Begin VB.TextBox txtLocate 
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   1200
            Width           =   4575
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   56
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select a Product"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdCloseLocate 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3600
         TabIndex        =   54
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Data dbLocate 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Data dbTrans 
      Caption         =   "Transaction"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame fraPOS 
      Height          =   6735
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdLocation 
         Caption         =   "Product Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4935
         MaskColor       =   &H8000000A&
         Picture         =   "frmPOS.frx":B7440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Receipt"
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
         Height          =   975
         Left            =   3960
         Picture         =   "frmPOS.frx":B774A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdTotal 
         Caption         =   "Total"
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
         Height          =   975
         Left            =   2925
         MaskColor       =   &H8000000A&
         Picture         =   "frmPOS.frx":B8014
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdSubtotal 
         Caption         =   "Subtotal"
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
         Height          =   975
         Left            =   2070
         MaskColor       =   &H8000000A&
         Picture         =   "frmPOS.frx":B8456
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdAmount 
         Caption         =   "Amount"
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
         Height          =   975
         Left            =   1215
         MaskColor       =   &H8000000A&
         Picture         =   "frmPOS.frx":B8898
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmPOS.frx":B8CDA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdClosePOS 
         Caption         =   "Close POS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6120
         MaskColor       =   &H8000000A&
         Picture         =   "frmPOS.frx":B99A4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4800
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   7080
         ScaleHeight     =   1665
         ScaleWidth      =   3945
         TabIndex        =   30
         Top             =   4800
         Width           =   3975
         Begin VB.Label lblDisplay 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   3
            Left            =   3075
            TabIndex        =   38
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label lblDisplay 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   2
            Left            =   3075
            TabIndex        =   37
            Top             =   840
            Width           =   765
         End
         Begin VB.Label lblDisplay 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   1
            Left            =   3075
            TabIndex        =   36
            Top             =   480
            Width           =   765
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Due"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   525
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tendered"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1290
         End
         Begin VB.Label lblDisplay 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   3075
            TabIndex        =   32
            Top             =   120
            Width           =   765
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   675
         End
      End
      Begin VB.TextBox txtPCode 
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   7830
         MaxLength       =   4
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   2040
      End
      Begin VB.TextBox txtPName 
         Height          =   375
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2040
         Width           =   4560
      End
      Begin VB.Data dbData 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDBGrid.DBGrid dbPOS 
         Bindings        =   "frmPOS.frx":B9DEE
         Height          =   2175
         Left            =   120
         OleObjectBlob   =   "frmPOS.frx":B9E03
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   10935
      End
      Begin VB.PictureBox picBlank 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2145
         ScaleWidth      =   10905
         TabIndex        =   46
         Top             =   2520
         Width           =   10935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Enter) - Amount"
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
         Left            =   120
         TabIndex        =   49
         Top             =   5950
         Width           =   1860
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=) - Total"
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
         Left            =   2520
         TabIndex        =   48
         Top             =   6250
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(+) - Subtotal"
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
         Left            =   2520
         TabIndex        =   47
         Top             =   5950
         Width           =   1440
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   10935
      End
      Begin VB.Label lblSubTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblPCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   18
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6810
         TabIndex        =   17
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8085
         TabIndex        =   16
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9735
         TabIndex        =   15
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label lblPName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3375
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SubTotal, Total, Amount, AmtChange, Due As Double

Private Sub cboProduct_Click()

If cboProduct.ListIndex = -1 Then Exit Sub

With dbLocate.Recordset

    .MoveFirst
    .FindFirst "[Brand Name]='" & cboProduct.Text & "'"
    
    txtLocate.Text = .Fields("Location")
    
End With

End Sub

Private Sub cboSearch_Click()

If cboSearch.ListIndex = -1 Then Exit Sub

txtSearch.Enabled = True
dbProd.Enabled = True
cmdInsert.Enabled = False

End Sub

Private Sub cmdAmount_Click()
    
If txtQty.Text = Empty Or Val(txtQty.Text) <= 0 Then
    MsgBox "Enter Quantity.", vbExclamation, "Required"
    txtQty.SetFocus
    Exit Sub
End If
              
If CheckOnHand(2, Val(txtQty.Text)) = "Low Stock" Then
    txtQty.Text = Empty
    txtQty.SetFocus
    Exit Sub
End If
      
txtAmount.Text = Val(txtQty.Text) * Val(txtPrice.Text)
        
txtAmount.SetFocus
    
lblSubTotal.Caption = "Amount"
lblTotal.Caption = Format(txtAmount.Text, "P ##,###,##0.00")

End Sub

Private Sub cmdCancel_Click()

'enable POS..
fraPOS.Enabled = True

'hide Amount..
picAmount.Visible = False

txtPCode.SetFocus

End Sub

Private Sub cmdClose_Click()

picForm.Visible = False
fraPOS.Enabled = True

End Sub

Private Sub cmdCloseLocate_Click()

picLocate.Visible = False
fraPOS.Enabled = True

End Sub

Private Sub cmdClosePOS_Click()

Unload Me

End Sub

Private Sub cmdInsert_Click()

With dbPrice.Recordset
    
    txtPCode.Text = ![Product Code]
    txtPName.Text = ![Brand Name]
    txtPrice.Text = !Price

End With

Call cmdClose_Click
txtQty.SetFocus

End Sub

Private Sub cmdLocation_Click()

cboProduct.ListIndex = -1
txtLocate.Text = Empty

picLocate.Visible = True

fraPOS.Enabled = False

End Sub

Private Sub cmdNew_Click()

ans = MsgBox("Create New Transaction?", _
    vbYesNo + vbQuestion, "Confirm")
    
If ans = vbNo Then Exit Sub

'reset all textbox, datas, etc..
Call ResetAll

'set focus to Product Code..
txtPCode.SetFocus

End Sub

Private Sub cmdOK_Click()

Amount = Val(txtTender.Text)
AmtChange = 0
Due = 0

If Amount >= Total Then
    AmtChange = Amount - Total
ElseIf Total > Amount Then
    Due = Amount - Total
End If

'display tender..
lblDisplay(1).Caption = Format(Amount, "P ##,###,##0.00")

'display change..
lblDisplay(3).Caption = Format(AmtChange, "P ##,###,##0.00")

'display due..
lblDisplay(2).Caption = Format(Due, "P ##,###,##0.00")

cmdPrint.Enabled = True

Call cmdCancel_Click

End Sub

Private Sub cmdPrint_Click()

'disable textbox and buttons..
cmdAmount.Enabled = False
cmdSubtotal.Enabled = False
cmdTotal.Enabled = False

txtPCode.Enabled = False
txtPName.Enabled = False
txtPrice.Enabled = False
txtQty.Enabled = False
txtAmount.Enabled = False

cmdPrint.Enabled = False

'write data to transaction table..
Call Transaction

With frmReceipt.RTFBox

    .Text = "ABIL PHARMACY STORE"
    .Text = .Text & vbCrLf
    .Text = .Text & String(35, "-")
    .Text = .Text & vbCrLf
    .Text = .Text & vbCrLf
    .Text = .Text & Date
    .Text = .Text & vbCrLf
    .Text = .Text & Time
    .Text = .Text & vbCrLf
    .Text = .Text & vbCrLf

    'get items..
    dbData.Recordset.MoveFirst

    Do While dbData.Recordset.EOF = False

        'get quantity..
        .Text = .Text & _
            Trim(Val(dbData.Recordset.Fields("Quantity")))

        'Tab..
        .Text = .Text & Space$(5)

        'get product name..
        .Text = .Text & _
            dbData.Recordset.Fields("Product Name")

        'next line..
        .Text = .Text & vbCrLf
        
        'Tab..
        .Text = .Text & Space$(6)
        
        'get amount..
        .Text = .Text & _
            Format(Val(dbData.Recordset.Fields("amount")), "P ##,###,##0.00")

        'Tab..
        .Text = .Text & Space$(7)

        'next line..
        .Text = .Text & vbCrLf

        dbData.Recordset.MoveNext

    Loop
    
    .Text = .Text & vbCrLf
    .Text = .Text & String(35, "-")

    'next line..
    .Text = .Text & vbCrLf

    'total..
    .Text = .Text & "Total:" & Space$(4) & Format(Total, "P ##,###,##0.00")

    'next line..
    .Text = .Text & vbCrLf

    'tendered..
    .Text = .Text & "Tendered: "
    .Text = .Text & Format(Amount, "P ##,###,##0.00")

    'next line..
    .Text = .Text & vbCrLf

    'due..
    .Text = .Text & "Due: "
    .Text = .Text & Space$(5)
    .Text = .Text & Format(Due, "P ##,###,##0.00")

    'next line..
    .Text = .Text & vbCrLf

    'change..
    .Text = .Text & "Change: "
    .Text = .Text & Space$(2)
    .Text = .Text & Format(AmtChange, "P ##,###,##0.00")
    
    .Text = .Text & vbCrLf
    .Text = .Text & vbCrLf
    .Text = .Text & "** Thank You **"
    
    .SelLength = 0
    .SelPrint Printer.hDC
    
End With

End Sub

Private Sub cmdSubtotal_Click()
    
Call AddToList
Call ClearBox
    
txtPCode.SetFocus
SubTotal = SubTotal + Val(txtAmount.Tag)
    
lblSubTotal.Caption = "SubTotal"
lblTotal.Caption = Format(SubTotal, "P ##,###,##0.00")

picBlank.Visible = False
dbPOS.Visible = True

txtAmount.Tag = 0

cmdSubtotal.Enabled = False

End Sub

Private Sub cmdTotal_Click()

If txtAmount.Tag <> Empty And _
    txtAmount.Tag <> 0 Then
        
        Call AddToList
        Call MinusProduct(txtPCode.Text, _
            Val(txtQty.Text))
End If

Call ClearBox

txtPCode.SetFocus
SubTotal = SubTotal + Val(txtAmount.Tag)

Total = SubTotal

lblSubTotal.Caption = "Total"
lblTotal.Caption = Format(Total, "P ##,###,##0.00")

picBlank.Visible = False
dbPOS.Visible = True

'display total..
lblDisplay(0).Caption = lblTotal.Caption

'disable POS..
fraPOS.Enabled = False

'show Amount form..
picAmount.Visible = True

txtTender.SetFocus

txtAmount.Tag = 0

cmdSubtotal.Enabled = False

End Sub

Private Sub dbProd_Click()

cmdInsert.Enabled = True

End Sub

Private Sub dbProd_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If CheckOnHand(0) = "Out of Stock" Then
    cmdInsert.Enabled = False
End If

Call CheckOnHand(1)

End Sub

Private Sub Form_Activate()

'hide toolbar..
mdiMain.Toolbar1.Visible = False

End Sub

Private Sub Form_Load()

Me.Left = 0
Me.Top = 0

'open temp table..
With dbData
    
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "temp"
    .Refresh
    
    With .Recordset
    
        'delete all records..
        Do While .EOF = False
        
            .Delete
            .MoveNext
        
        Loop
    
    End With
    
End With

dbPOS.Refresh
dbPOS.ReBind

'open product table..
With dbPrice
    
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = _
        "SELECT [Product Code], [Generic Name], [Brand Name], " & _
        "Price, OnHand, [Reorder Level] FROM Product ORDER BY [Product Code], [Generic Name], [Brand Name]"
    .Refresh
    
End With

'open product table..
With dbLocate
    
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "Product"
    .Refresh

End With

'open transaction table..
With dbTrans
    
    .DatabaseName = App.Path & "\database.mdb"
    .RecordSource = "Transaction"
    .Refresh

End With

'set True to AutoSize of PicForm..
picForm.AutoSize = True

'set True to AutoSize of picAmount..
picAmount.AutoSize = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

'show toolbar..
mdiMain.Toolbar1.Visible = True

End Sub

Private Sub txtAmount_Change()

If txtAmount.Tag = Empty Then
    txtAmount.Tag = txtAmount.Text
End If

txtAmount.Text = Format(txtAmount.Text, "##,###,##0.00")

End Sub

Private Sub txtAmount_GotFocus()

If txtPCode.Text = Empty Then
    txtPCode.SetFocus
    Exit Sub
End If

If txtQty.Text = Empty Or Val(txtQty.Text) <= 0 Then
    txtQty.SetFocus
    Exit Sub
End If

cmdAmount.Enabled = False
cmdSubtotal.Enabled = True
cmdTotal.Enabled = True

End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)

'get subtotal..
If Chr(KeyAscii) = "+" Then
    
    Call MinusProduct(txtPCode.Text, Val(txtQty.Text))
    
    Call AddToList
    Call ClearBox
    
    txtPCode.SetFocus
    SubTotal = SubTotal + Val(txtAmount.Tag)
    
    lblSubTotal.Caption = "SubTotal"
    lblTotal.Caption = Format(SubTotal, "P ##,###,##0.00")

    picBlank.Visible = False
    dbPOS.Visible = True
    
    cmdSubtotal.Enabled = False
    
    txtAmount.Tag = 0

'get total..
ElseIf Chr(KeyAscii) = "=" Then
    
    If txtAmount.Tag <> Empty Then
        Call MinusProduct(txtPCode.Text, Val(txtQty.Text))
    End If
    
    If txtAmount.Text <> Empty Then
        
        Call AddToList
        
    End If
    
    Call ClearBox
    
    txtPCode.SetFocus
    SubTotal = SubTotal + Val(txtAmount.Tag)
    
    Total = SubTotal
    
    lblSubTotal.Caption = "Total"
    lblTotal.Caption = Format(Total, "P ##,###,##0.00")

    picBlank.Visible = False
    dbPOS.Visible = True
    cmdSubtotal.Enabled = False
    
    'display total..
    lblDisplay(0).Caption = lblTotal.Caption
    
    'disable POS..
    fraPOS.Enabled = False
    
    'show Amount form..
    picAmount.Visible = True
    
    txtTender.SetFocus
    
    txtAmount.Tag = 0
    
End If

End Sub

Private Sub txtPCode_DblClick()

'disable POS form..
fraPOS.Enabled = False

'show search form..
picForm.Visible = True

'disable dbProd DBGrid..
'dbProd.Refresh

'disable txtSearch..
txtSearch.Enabled = False

'disable Insert button..
cmdInsert.Enabled = False

'unselect cboSearch..
cboSearch.ListIndex = -1

'set focus to cboSearch..
cboSearch.SetFocus

End Sub

Private Sub txtPCode_GotFocus()

If txtPCode.Text <> Empty And txtQty.Text = Empty Then
    txtQty.SetFocus
ElseIf txtPCode.Text <> Empty And txtQty.Text <> Empty Then
    txtAmount.SetFocus
End If

End Sub

Private Sub txtPCode_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = "=" Then

    'erase equal sign..
    KeyAscii = 0
    
    If lblTotal.Caption = "P0.00" Then Exit Sub
    
    Call ClearBox
    
    If txtAmount.Text <> Empty Then
        
        Call AddToList
        
    End If
    
    txtPCode.SetFocus
    SubTotal = SubTotal + Val(txtAmount.Tag)
    
    Total = SubTotal
    
    lblSubTotal.Caption = "Total"
    lblTotal.Caption = Format(Total, "P ##,###,##0.00")

    picBlank.Visible = False
    dbPOS.Visible = True
    
    'display total..
    lblDisplay(0).Caption = lblTotal.Caption
    
    'disable POS..
    fraPOS.Enabled = False
    
    'show Amount form..
    picAmount.Visible = True
    
    txtTender.SetFocus
    
End If

Select Case KeyAscii

    Case 65 To 90       'A to Z..
    Case 97 To 122
        KeyAscii = KeyAscii - 32
    Case 48 To 57       '0-9
    Case vbKeyBack
    Case vbKeyReturn
        Call SearchPCode
    Case Else
        KeyAscii = 0

End Select

End Sub

Private Sub txtPName_GotFocus()

If txtPName.Text = Empty Then
    txtPCode.SetFocus
Else
    txtQty.SetFocus
End If

End Sub

Private Sub txtPrice_Change()

txtPrice.Text = Format(txtPrice.Text, "##,###,##0.00")

End Sub

Private Sub txtPrice_GotFocus()

If txtPrice.Text = Empty Then
    txtPCode.SetFocus
Else
    txtQty.SetFocus
End If

End Sub

Private Sub txtQty_GotFocus()

If txtPCode.Text = Empty Then
    
    txtPCode.SetFocus
    Exit Sub
   
End If
            
If CheckOnHand(2, Val(txtQty.Text)) = "Low Stock" Then
    txtQty.Text = Empty
    Exit Sub
End If

txtAmount.Tag = Empty

cmdAmount.Enabled = True
cmdSubtotal.Enabled = False
cmdTotal.Enabled = False

End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 48 To 57       '0-9
    Case vbKeyBack
    
    Case vbKeyReturn
            
        If CheckOnHand(2, Val(txtQty.Text)) = "Low Stock" Then
            txtQty.Text = Empty
            Exit Sub
        End If

        If txtQty.Text = Empty Or Val(txtQty.Text) <= 0 Then
            MsgBox "Enter Quantity.", vbExclamation, "Required"
            txtQty.SetFocus
            Exit Sub
        End If
        
        txtAmount.Text = Val(txtQty.Text) * Val(txtPrice.Text)
        
        txtAmount.SetFocus
    
        lblSubTotal.Caption = "Amount"
        lblTotal.Caption = Format(txtAmount.Text, "P ##,###,##0.00")
    
    Case Else
        KeyAscii = 0

End Select

End Sub

Private Sub txtSearch_Change()

If txtSearch.Text = "" Then Exit Sub

With dbPrice.Recordset

    .MoveFirst
    
    Select Case cboSearch.ListIndex
    
        'by product code..
        Case 0
            .FindFirst "[Product Code] like '" & txtSearch.Text & "*'"
            
        'by brand name..
        Case 1
            .FindFirst "[Brand Name] like '" & txtSearch.Text & "*'"
            
        'by generic name..
        Case 2
            .FindFirst "[Generic Name] like '" & txtSearch.Text & "*'"
            
    End Select
    
    If .NoMatch = True Then
        MsgBox "Product not found.", vbExclamation, "Search"
        txtSearch.Text = Empty
        Exit Sub
    End If
    
    cmdInsert.Enabled = False
    
End With

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45, 47         'slash and hyphen
    Case 65 To 90       'A to Z..
    Case 97 To 122
    Case 48 To 57       '0-9
    Case vbKeyBack
    Case Else
        KeyAscii = 0

End Select

End Sub

Sub SearchPCode()

If txtPCode.Text = Empty Then Exit Sub

With dbPrice.Recordset

    .MoveFirst
    .FindFirst "[Product Code] = '" & txtPCode.Text & "'"
    
    If .NoMatch = True Then
        MsgBox "Product not found.", vbExclamation, "Search"
        txtPCode.Text = Empty
        txtPCode.SetFocus
        Exit Sub
    End If
    
    txtPCode.Text = ![Product Code]
    txtPName.Text = ![Brand Name]
    txtPrice.Text = !Price
    
    txtQty.SetFocus
    
End With

End Sub

Sub ClearBox()

txtPCode.Text = Empty
txtPName.Text = Empty
txtPrice.Text = Empty
txtQty.Text = Empty
txtAmount.Text = Empty

End Sub

Sub AddToList()

With dbData.Recordset

    .AddNew
    .Fields(0) = txtPCode.Text
    .Fields(1) = txtPName.Text
    .Fields(2) = txtPrice.Text
    .Fields(3) = txtQty.Text
    .Fields(4) = 0
    .Fields(5) = txtAmount.Text
    .Fields(6) = Month(Now)
    .Fields(7) = Year(Now)
    .Update
    
End With

dbData.Refresh
dbPOS.Refresh

cboProduct.AddItem txtPName.Text

End Sub

Private Sub txtTender_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 46, 48 To 57      'period (.), 0-9
    Case vbKeyBack
    Case vbKeyReturn
        Call cmdOK_Click
    Case Else
        KeyAscii = 0
End Select

End Sub

Sub ResetAll()

'open temp table..
With dbData
    With .Recordset
    
        'if there is a transaction..
        If .EOF = False Or .RecordCount > 0 Then
        
            .MoveFirst
        
            'delete all records..
            Do While .EOF = False
            
                .Delete
                .MoveNext
            
            Loop
    
        End If
    
    End With

End With

dbData.Refresh
dbPrice.Refresh

dbProd.Refresh

dbPOS.Refresh
dbPOS.ReBind

'dbPOS.Visible = False
'picBlank.Visible = True

SubTotal = 0
Total = 0

txtTender.Text = Empty

lblTotal.Caption = "P0.00"
lblSubTotal.Caption = "Amount"

lblDisplay(0).Caption = "P0.00"
lblDisplay(1).Caption = "P0.00"
lblDisplay(2).Caption = "P0.00"
lblDisplay(3).Caption = "P0.00"

Call ClearBox

cmdAmount.Enabled = False
cmdSubtotal.Enabled = False
cmdTotal.Enabled = False

cmdPrint.Enabled = False

txtPCode.Enabled = True
txtQty.Enabled = True
txtAmount.Enabled = True
txtPName.Enabled = True
txtPrice.Enabled = True

cboProduct.Clear
txtLocate.Text = Empty

End Sub

Sub Transaction()

With dbTrans.Recordset

    'check if user has a present transaction..
    If dbData.Recordset.EOF = True And _
        dbData.Recordset.RecordCount = 0 Then
            Exit Sub
    End If
    
    'if yes, write data to Transaction table..
    dbData.Recordset.MoveFirst

    Do While dbData.Recordset.EOF = False
    
        .AddNew
        
        .Fields("Product Code") = _
            dbData.Recordset.Fields("Product Code")
        .Fields("Product Name") = _
            dbData.Recordset.Fields("Product Name")
        .Fields("Quantity") = _
            dbData.Recordset.Fields("Quantity")
        .Fields("Price") = _
            dbData.Recordset.Fields("Price")
        .Fields("Date") = Date
        
        .Update
        
        dbData.Recordset.MoveNext
        
    Loop

End With

End Sub

Function CheckOnHand(ByVal Checking As Integer, Optional ByVal Num As Integer) As String

With dbPrice.Recordset

    Select Case Checking

        'check OnHand..
        Case 0
            If .Fields("OnHand") = 0 Then
                MsgBox "'" & .Fields("Brand Name") & "' is out of stock.", _
                    vbExclamation, "Message"
                CheckOnHand = "Out of Stock"
            Else
                CheckOnHand = "OK"
            End If
        
        'compare OnHand and ReOrder level..
        Case 1
            If .Fields("OnHand") < _
                .Fields("ReOrder Level") Then
                
                MsgBox "'" & .Fields("Brand Name") & "' is below the ReOrder Level.", _
                    vbExclamation, "Message"
                CheckOnHand = "Below ReOrder Level"
            Else
                CheckOnHand = "OK"
            End If

        'compare OnHand and Quantity..
        Case 2
            If .Fields("OnHand") < Num Then
                MsgBox "The available stocks for '" & .Fields("Brand Name") & _
                    "' is " & .Fields("OnHand"), _
                    vbExclamation, "Message"
                CheckOnHand = "Low Stock"
            Else
                CheckOnHand = "OK"
            End If
            
    
    End Select

End With

End Function

Sub MinusProduct(ByVal ProdCode As String, ByVal Qty As Integer)

With dbPrice.Recordset
    
    .MoveFirst
    .FindFirst "[Product Code]='" & ProdCode & "'"
    
    .Edit
    .Fields("OnHand") = .Fields("OnHand") - Qty
    .Update
    
End With

End Sub
