VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form posterminal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "EMMS"
   ClientHeight    =   10260
   ClientLeft      =   -345
   ClientTop       =   0
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   14475
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picmsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   600
      ScaleHeight     =   705
      ScaleWidth      =   13185
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   13215
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   240
      End
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ENTER Key For New Transaction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   13215
      End
   End
   Begin VB.PictureBox picv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   5760
      ScaleHeight     =   1305
      ScaleWidth      =   2985
      TabIndex        =   21
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   240
         PasswordChar    =   "l"
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC to Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor Pass:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox picq 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   5760
      ScaleHeight     =   1305
      ScaleWidth      =   2985
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Quantity:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC to Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   7680
      Width           =   2895
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   9360
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   1058
      ButtonWidth     =   3493
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F2 - Void   "
            Key             =   "Void"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F3 - &Look Up   "
            Key             =   "Look Up"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F4 - &Sale Type    "
            Key             =   "Sale Type"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F6 - &Cash Tend   "
            Key             =   "Cash Tend"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F10 - &Log Off  "
            Key             =   "Switch User"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F11 - &Exit  "
            Key             =   "Exit"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   960
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":1992
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":3324
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":4CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":6648
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":7FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":996C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":B2FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":CC90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":E624
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":F300
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":FBE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":108BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":11598
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":12274
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":12F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "posterminal.frx":13C2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   9960
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
            Picture         =   "posterminal.frx":14508
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Cashier:"
            TextSave        =   "Cashier:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11977
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "posterminal.frx":148A4
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10/9/2005"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2:28 AM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox piclookup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   3360
      ScaleHeight     =   3945
      ScaleWidth      =   7905
      TabIndex        =   36
      Top             =   2280
      Visible         =   0   'False
      Width           =   7935
      Begin MSComctlLib.ListView lvsearch 
         Height          =   2775
         Left            =   240
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Description"
            Object.Width           =   10231
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "UPC"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtsearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Part Description then press ENTER. Press ESC to Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   40
         Top             =   3600
         Width           =   7335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Desc:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   11
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.PictureBox piqs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   5640
      ScaleHeight     =   1545
      ScaleWidth      =   3225
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "REGULAR SALE"
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC to Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Sale Type:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.PictureBox piccash 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   5760
      ScaleHeight     =   1305
      ScaleWidth      =   2985
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC to Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   10
         Left            =   360
         TabIndex        =   32
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Cash:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   31
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   5760
      ScaleHeight     =   1305
      ScaleWidth      =   2985
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "ALL"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Option:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC to Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   27
         Top             =   960
         Width           =   2535
      End
   End
   Begin MSComctlLib.ListView lstsale 
      Height          =   6015
      Left            =   600
      TabIndex        =   41
      Top             =   1440
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   10610
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UPC"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "LONG DESCRIPTION"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SHORT DESCRIPTION"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "UNIT PRICE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "QUANTITY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "TOTAL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "SUP ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CAT ID"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblsaleid 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "REGULAR SALE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   9840
      TabIndex        =   35
      Top             =   960
      Width           =   3975
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   600
      Picture         =   "posterminal.frx":14C3E
      Top             =   8760
      Width           =   480
   End
   Begin VB.Label lblsaletype 
      BackStyle       =   0  'Transparent
      Caption         =   "REGULAR SALE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label LBLTIP 
      BackStyle       =   0  'Transparent
      Caption         =   "Press <F12> to Enter quantity. Cashier should enter item quantity before entering or scanning UPC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1200
      TabIndex        =   11
      Top             =   8760
      Width           =   5655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "UPC/Barcode:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   7680
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   8400
      TabIndex        =   9
      Top             =   8520
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Tendered:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   8040
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sales:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   0
      Left            =   8400
      TabIndex        =   7
      Top             =   7560
      Width           =   3255
   End
   Begin VB.Label LBLCHANGE 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   10560
      TabIndex        =   6
      Top             =   8520
      Width           =   3255
   End
   Begin VB.Label LBLCASHTEND 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   10560
      TabIndex        =   5
      Top             =   8040
      Width           =   3255
   End
   Begin VB.Label LBLTOTAL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   7560
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Point of Sales Terminal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9240
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   9945
      Left            =   0
      Picture         =   "posterminal.frx":15080
      Stretch         =   -1  'True
      Top             =   720
      Width           =   14475
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   930
      Left            =   0
      Picture         =   "posterminal.frx":17BAB
      Stretch         =   -1  'True
      ToolTipText     =   "For School Project Purpose Only, Not for Sale..."
      Top             =   0
      Width           =   15405
   End
End
Attribute VB_Name = "posterminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intquantity As Long
Dim inttotal As Currency
Dim intcash As Currency
Dim intchange As Currency
Dim iloop As Byte
Dim s As String
Dim ints As Integer
Dim ilst As Integer
Dim lstid, lstupc, lstlong, lstshort, lstunitprice, lstquantity, lsttotal, lstsup, lstcat As String
Dim sdesc, supc As String

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        txt(0).Enabled = True
        Toolbar1.Enabled = True
        lstsale.Enabled = True
        txt(0).SetFocus
        piqs.Visible = False
End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lblsaletype.Caption = Combo1.Text
        txt(0).Enabled = True
        Toolbar1.Enabled = True
        lstsale.Enabled = True
        piqs.Visible = False
        txt(0).SetFocus
    End If
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyEscape
                'txt(1) = "1"
                pic.Visible = False
                txt(0).Enabled = True
                Toolbar1.Enabled = True
                lstsale.Enabled = True
                txt(0).SetFocus
        End Select
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        picv.Visible = True
        txt(2).SetFocus
        pic.Visible = False
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Image1.Width = Me.Width
    
    Combo1.AddItem "REGULAR SALE"
    Combo1.AddItem "WHOLESALE"
    Combo1.AddItem "CREDIT SALE"
    
    Combo2.AddItem "ALL"
    Combo2.AddItem "SELECTED ITEM"
    
    Set ac = New ADODB.Connection
    Set ar = New ADODB.Recordset
    
    Call dbconek
    ac.Open strConek
    
    With ar
        criteria = "Select *From sale"
        .Open criteria, strConek, adOpenStatic, adLockOptimistic
            .MoveLast
                lblsaleid.Caption = !saleid + 1
        .Close
    End With
    
    StatusBar2.Panels(3) = currentemmsuser
    
End Sub

Private Sub Image6_Click()
Load frmMMSPurchase
frmMMSPurchase.Show vbModal
End Sub
Private Sub Image8_Click()
    Unload Me
End Sub

Private Sub lstsale_KeyPress(KeyAscii As Integer)
Dim retim As Integer
    If KeyAscii = 13 Then
            pic.Visible = True
            Combo2.SetFocus
            lstsale.Enabled = False
                txt(0).Enabled = False
                Toolbar1.Enabled = False
            
    End If
End Sub





Private Sub lvsearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
        piclookup.Visible = False
        txt(0).Enabled = True
        txt(0).SetFocus
        lstsale.Enabled = True
        Toolbar1.Enabled = True
    End If
End Sub

Private Sub lvsearch_KeyPress(KeyAscii As Integer)
    txt(0) = lvsearch.ListItems(lvsearch.SelectedItem.Index).SubItems(1)
    txtsearch = ""
    lvsearch.ListItems.Clear
    piclookup.Visible = False
    txt(0).Enabled = True
    lstsale.Enabled = True
    Toolbar1.Enabled = True
    Call txt_KeyPress(0, 13)
    txt(0).SetFocus
End Sub

Private Sub picmsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        picmsg.Visible = False
        Timer1.Enabled = False
        txt(0).Enabled = True
        Toolbar1.Enabled = True
        lstsale.Enabled = True
        lstsale.ListItems.Clear
        
        'set the values to 0
        inttotal = 0
        intcash = 0
        intchange = 0
        intquantity = 0
        
        'LBLTOTAL.Caption = "0.00"
        'LBLCASHTEND.Caption = "0.00"
        'LBLCHANGE.Caption = "0.00"
        'end
        'txt(0).SetFocus
        Me.Refresh
        
        Unload Me
        Load posterminal
        posterminal.Show
        
    End If
End Sub

Private Sub Timer1_Timer()
    iloop = iloop + 1
    If iloop = 1 Then lblmsg.ForeColor = &HFFFFC0
    If iloop = 2 Then lblmsg.ForeColor = vbBlack
    If iloop = 2 Then iloop = 0
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    txt(0).SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case Index
    Case 0
        Select Case KeyCode
            Case vbKeyF2
                If lstsale.ListItems.Count > 0 Then
                    lstsale.SetFocus
                    LBLTIP.Caption = "Use arrow key to select item to void. Press ENTER key to Void All"
                Else
                    MsgBox "There are no items to void", vbCritical, "Cashier Error"
                    txt(0).SetFocus
                    Exit Sub
                End If
            Case vbKeyF3
                Dim intitem As ListItem
                txt(0).Enabled = False
                Toolbar1.Enabled = False
                piclookup.Visible = True
                txtsearch.SetFocus
                SendKeys "{HOME}+{END}"
            Case vbKeyF4
                If lstsale.ListItems.Count = 0 Then
                    piqs.Visible = True
                    txt(0).Enabled = False
                    Toolbar1.Enabled = False
                    lstsale.Enabled = False
                    Combo1.SetFocus
                Else
                    MsgBox "You have already enter an item for this type of sale. Void the transaction if necessary.", vbCritical, "Cashier Error"
                    txt(0).SetFocus
                    Exit Sub
                End If
            Case vbKeyF6
                If lstsale.ListItems.Count = 0 Then
                    MsgBox "There are no item(s) bought", vbCritical, "Cashier Error"
                    txt(0).SetFocus
                    Exit Sub
                Else
                    piccash.Visible = True
                    txt(3).SetFocus
                    txt(0).Enabled = False
                    Toolbar1.Enabled = False
                    lstsale.Enabled = False
                End If
            Case vbKeyF10
                Load frmxLogIn
                frmxLogIn.Show vbModal
            Case vbKeyF11
                Unload Me
            Case vbKeyF12
                    picq.Visible = True
                    txt(1).SetFocus
                    SendKeys "{home}+{end}"
                    txt(0).Enabled = False
                    Toolbar1.Enabled = False
                    lstsale.Enabled = False
        End Select
    Case 1
        Select Case KeyCode
            Case vbKeyEscape
                txt(1) = "1"
                picq.Visible = False
                txt(0).Enabled = True
                Toolbar1.Enabled = True
                lstsale.Enabled = True
                txt(0).SetFocus
        End Select
    Case 2
        Select Case KeyCode
            Case vbKeyEscape
                'txt(1) = "1"
                picv.Visible = False
                txt(0).Enabled = True
                Toolbar1.Enabled = True
                lstsale.Enabled = True
                txt(0).SetFocus
        End Select
    Case 3
        Select Case KeyCode
            Case vbKeyEscape
                'txt(1) = "1"
                piccash.Visible = False
                txt(0).Enabled = True
                Toolbar1.Enabled = True
                lstsale.Enabled = True
                txt(0).SetFocus
        End Select
End Select
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Select Case Index
    Case 0
        '*****************************************************
        'For different sales type
        If KeyAscii = 13 Then
            'inttotal = 0
            
            Set ac = New ADODB.Connection
            Set ar = New ADODB.Recordset
        
            Call dbconek
            ac.Open strConek
            
            'Getting the discount, credit and regular amount from the database
            With ar
                criteria = "Select *From itemmaster Where upc='" & txt(0) & "'"
                .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    If .RecordCount >= 1 Then
                        Select Case lblsaletype.Caption
                            Case "REGULAR SALE"
                                
                                lstsale.ListItems.Add , , txt(0)
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(1) = !longdesc
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(2) = !shortdesc
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(3) = Format(CCur(!unitprice), "###,###,##0.00")
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(4) = txt(1)
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(5) = Format(CCur(!unitprice * txt(1)), "###,###,##0.00")
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(6) = !supplierid
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(7) = !categoryid
                                'lstsale.AddItem padspace(txt(0), 15) & padspace(!longdesc, 50) & padspace(!shortdesc, 25) & padspace1(Format(CCur(!unitprice), "###,###,##0.00"), 10) & padspace1("@", 3) & padspace1(txt(1), 10) & padspace1(Format(CCur(!unitprice * txt(1)), "###,###,##0.00"), 15) & padspace(!supplierid, 10) & padspace(!categoryid, 10)
                                'lstsale.AddItem padc(txt(0), 15) & " " & padc(Trim(!longdesc), 60) & padspace(!shortdesc, 50) & padspace1(Format(!unitprice, "###,###,##0.00"), 20) & padspace1(txt(1), 20) & padspace1(Format(!unitprice * Val(txt(1)), "###,###,##0.00"), 20) & padspace1(!supplierid, 10) & padspace1(!categoryid, 10)
                                inttotal = Val(txt(1)) * !unitprice
                                txt(1) = "1"
                                txt(0) = ""
                                LBLTOTAL.Caption = Format(CCur(LBLTOTAL.Caption) + CCur(inttotal), "###,###,##0.00")
                            Case "WHOLESALE"
                                lstsale.ListItems.Add , , txt(0)
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(1) = !longdesc
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(2) = !shortdesc
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(3) = Format(CCur(!unitprice), "###,###,##0.00")
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(4) = txt(1)
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(5) = Format(CCur((!unitprice * txt(1)) / (1 + !discount)), "###,###,##0.00")
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(6) = !supplierid
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(7) = !categoryid
                                
                                'lstsale.AddItem addspace(txt(0), 15) & addspace(!longdesc, 50) & addspace(!shortdesc, 25) & addspace(Format(CCur(!unitprice), "###,###,##0.00"), 10) & addspace("@", 3) & addspace(txt(1), 10) & addspace(Format(CCur((!unitprice * txt(1)) / (1 + !discount)), "###,###,##0.00"), 15) & addspace(!supplierid, 10) & addspace(!categoryid, 10)
                                inttotal = CCur(txt(1)) * CCur(!unitprice) / (1 + !discount)
                                txt(1) = "1"
                                txt(0) = ""
                                LBLTOTAL.Caption = Format(CCur(LBLTOTAL.Caption) + CCur(inttotal), "###,###,##0.00")
                            Case "CREDIT SALE"
                                lstsale.ListItems.Add , , txt(0)
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(1) = !longdesc
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(2) = !shortdesc
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(3) = Format(CCur(!unitprice), "###,###,##0.00")
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(4) = txt(1)
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(5) = Format(CCur((!unitprice * txt(1)) * (1 + !interestincredit)), "###,###,##0.00")
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(6) = !supplierid
                                lstsale.ListItems(lstsale.ListItems.Count).SubItems(7) = !categoryid
                                
                                'lstsale.AddItem addspace(txt(0), 15) & addspace(!longdesc, 50) & addspace(!shortdesc, 25) & addspace(Format(CCur(!unitprice), "###,###,##0.00"), 10) & addspace("@", 3) & addspace(txt(1), 10) & addspace(Format(CCur((!unitprice * txt(1)) * (1 + !interestincredit)), "###,###,##0.00"), 15) & addspace(!supplierid, 10) & addspace(!categoryid, 10)
                                inttotal = CCur(txt(1)) * CCur(!unitprice) * (1 + !interestincredit)
                                txt(1) = "1"
                                txt(0) = ""
                                LBLTOTAL.Caption = Format(CCur(LBLTOTAL.Caption) + CCur(inttotal), "###,###,##0.00")
                        End Select
                    Else
                        txt(0).SetFocus
                        SendKeys "{home}+{end}"
                        Exit Sub
                    End If
                .Close
            End With
            'snd
        End If
        'end
    Case 1 'For quantity
        If KeyAscii = 13 Then
            'intquantity = txt(1)
            txt(0).Enabled = True
            Toolbar1.Enabled = True
            lstsale.Enabled = True
            txt(0).SetFocus
            picq.Visible = False
        End If
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    Case 2 'For Password check
        If KeyAscii = 13 Then
            Select Case Combo2.Text
                Case "ALL"
                    Set ac = New ADODB.Connection
                    Set ar = New ADODB.Recordset
        
                    Call dbconek
                    ac.Open strConek
        
                    With ar
                        criteria = "Select *From employeemaster Where password='" & txt(2) & "'"
                        .Open criteria, strConek, adOpenStatic, adLockOptimistic
                            If .RecordCount >= 1 Then
                                If !Position = "MANAGER" Or !Position = "SUPERVISOR" Then
                                    lstsale.ListItems.Clear
                                    txt(0).Enabled = True
                                    txt(0).SetFocus
                                    lstsale.Enabled = True
                                    Toolbar1.Enabled = True
                                    picv.Visible = False
                                    inttotal = 0
                                    LBLTOTAL.Caption = "0.00"
                                Else
                                    MsgBox "The password you entered is not allowed to void an item", vbCritical, "Void Error"
                                    txt(2).SetFocus
                                    SendKeys "{home}+{end}"
                                End If
                            Else
                                MsgBox "Invalid supervisor/manager password", vbCritical, "Void Error"
                                txt(2).SetFocus
                                SendKeys "{home}+{end}"
                            End If
                        .Close
                    End With
                Case "SELECTED ITEM"
                    Dim i As Integer
                    Dim minusamount As String
                    
                    Set ac = New ADODB.Connection
                    Set ar = New ADODB.Recordset
        
                    Call dbconek
                    ac.Open strConek
                    
                    
                    
                    With ar
                        criteria = "Select *From employeemaster Where password='" & txt(2) & "'"
                        .Open criteria, strConek, adOpenStatic, adLockOptimistic
                            If .RecordCount >= 1 Then
                                
                                If !joblevel = "1" Or !joblevel = 2 Then
                                
                                    minusamount = lstsale.ListItems(lstsale.SelectedItem.Index).SubItems(5)
                                    LBLTOTAL.Caption = Format(CCur(LBLTOTAL.Caption) - minusamount, "###,###,##0.00")
                                                                        
                                    lstsale.ListItems.Remove (lstsale.SelectedItem.Index)
                                    txt(0).Enabled = True
                                    txt(0).SetFocus
                                    lstsale.Enabled = True
                                    Toolbar1.Enabled = True
                                    picv.Visible = False
                                Else
                                    MsgBox "The password you entered is not allowed to void an item", vbCritical, "Void Error"
                                    txt(2).SetFocus
                                    SendKeys "{home}+{end}"
                                End If
                            Else
                                MsgBox "Invalid supervisor/manager password", vbCritical, "Void Error"
                                txt(2).SetFocus
                                SendKeys "{home}+{end}"
                            End If
                        .Close
                    End With
            End Select
        End If
    Case 3 'cash
        If KeyAscii = 13 Then
            If txt(3) <= LBLTOTAL - 1 Then
                txt(3).SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            Else
                'Calculate
                '************************************************************'
                intcash = txt(3).Text
                intchange = CCur(intcash) - LBLTOTAL.Caption
                LBLCASHTEND.Caption = Format(CCur(intcash), "###,###,##0.00")
                LBLCHANGE.Caption = Format(CCur(intchange), "###,###,##0.00")
            
                piccash.Visible = False
                picmsg.Visible = True
                
                
                
                'Save to parent table "sale"
                '***********************************************************'
                
                Set ac = New ADODB.Connection
                Set ar = New ADODB.Recordset
                
                Call dbconek
                ac.Open strConek
                                
                With ar
                    criteria = "Select *From sale"
                    .Open criteria, strConek, adOpenStatic, adLockOptimistic
                        .AddNew
                            !saleid = lblsaleid
                            !Date = Format(Date, "mm/dd/yyyy")
                            !employee = currentemmsuser
                            !salestotal = LBLTOTAL
                        .Update
                    .Close
                End With
                '**********************************************************
                                
                For ilst = 1 To lstsale.ListItems.Count
                    lstupc = lstsale.ListItems(ilst).Text
                    lstlong = lstsale.ListItems(ilst).SubItems(1)
                    lstshort = lstsale.ListItems(ilst).SubItems(2)
                    lstunitprice = lstsale.ListItems(ilst).SubItems(3)
                    lstquantity = lstsale.ListItems(ilst).SubItems(4)
                    lsttotal = lstsale.ListItems(ilst).SubItems(5)
                    lstsup = lstsale.ListItems(ilst).SubItems(6)
                    lstcat = lstsale.ListItems(ilst).SubItems(7)
                    
                    'save to child table saledetails
                    '***********************************************************
                    
                    Set ac = New ADODB.Connection
                    Set ar = New ADODB.Recordset
                    
                    Call dbconek
                    ac.Open strConek
                    
                    With ar
                        criteria = "Select *From saledetails"
                        .Open criteria, strConek, adOpenStatic, adLockOptimistic
                            .AddNew
                                !saleid = lblsaleid.Caption
                                !upc = lstupc
                                !longdesc = lstlong
                                !shortdesc = lstshort
                                !unitprice = lstunitprice
                                !quantity = lstquantity
                                !total = lsttotal
                                !supplierid = lstsup
                                !categoryid = lstcat
                                !Date = Format(Date, "mm/dd/yyyy")
                            .Update
                        .Close
                    End With
                    '***********************************************************
                    
                    'Deduct to inventory
                    '***********************************************************
                    Set ac = New ADODB.Connection
                    Set ar = New ADODB.Recordset
                    
                    Call dbconek
                    ac.Open strConek
                    
                    With ar
                        criteria = "Select *From itemmaster Where upc='" & lstupc & "'"
                        .Open criteria, strConek, adOpenStatic, adLockOptimistic
                            !stock = !stock - lstquantity
                            .Update
                        .Close
                    End With
                    '***********************************************************
                Next
                
                Dim longname, lname As Long
                                
                Dim ii As Integer
                    
                'Printer.Print Space(longname - Len("Planet Sports")) & "Planet Sports"
                'Printer.Print "LCC Branch Felix Plazo Naga City"
                'Printer.Print Space(longname - Len("Tin: 465-888-0987")) & "Tin: 465-888-0987"
                'Printer.Print " "
                'Printer.Print " "
            
                
                'For ii = 1 To lstsale.ListItems.Count
                '    Printer.Print lstsale.ListItems(ii).SubItems(2) & Space(28) & lstsale.ListItems(ii).SubItems(4) & Space(41) & lstsale.ListItems(ii).SubItems(5)
                'Next
                'Printer.Print ""
                'Printer.Print ""
                'Printer.Print "TOTAL : " & Space(lname - Len("TOTAL : ")) & LBLTOTAL
                'Printer.Print "CASH  : " & Space(lname - Len("CASH  : ")) & LBLCASHTEND
                'Printer.Print "CHANGE: " & Space(lname - Len("CHANGE: ")) & LBLCHANGE
                'Printer.Print ""
                'Printer.Print Space(50 - Len("This is your oficial receipt")) & "This is your oficial receipt"
                'Printer.Print Space(50 - Len("Thank You!!! Come Again")) & "Thank You!!! Come Again"
                
                'Dim longname As Long
                
                longname = Len(lstsale.ListItems(1).SubItems(2) & Space(30 - Len(lstsale.ListItems(1).SubItems(2))) & lstsale.ListItems(1).SubItems(4) & Space(15 - Len(lstsale.ListItems(1).SubItems(4))) & lstsale.ListItems(1).SubItems(5))
                
                Open App.Path & "\Reports\Purchase.txt" For Append As #1
                    Print #1, "Store Name: Planet Sports"
                    Print #1, "Branch Name: LCC Branch"
                    Print #1, "Address: Felix Plazo Naga City"
                    Print #1, "Tin: TNW-001-0987-67"
                    Print #1, "Transaction Number: " & lblsaleid
                    Print #1,
                    Print #1, "Item Description" & Space(30 - Len("Item Description")) & "Quantity" & Space(15 - Len("Quantity")) & "Total"
                    Print #1,
                    For ii = 1 To lstsale.ListItems.Count
                        Print #1, lstsale.ListItems(ii).SubItems(2) & Space(30 - Len(lstsale.ListItems(ii).SubItems(2))) & lstsale.ListItems(ii).SubItems(4) & Space(15 - Len(lstsale.ListItems(ii).SubItems(4))) & lstsale.ListItems(ii).SubItems(5)
                    Next
                    Print #1,
                    Print #1,
                    Print #1, Space(longname - Len("TOTAL   : " & LBLTOTAL.Caption)) & "TOTAL   : " & LBLTOTAL.Caption
                    Print #1, Space(longname - Len("CASHTEND: " & LBLCASHTEND.Caption)) & "CASHTEND: " & LBLCASHTEND.Caption
                    Print #1, Space(longname - Len("CHANGE  : " & LBLCHANGE.Caption)) & "CHANGE  : " & LBLCHANGE.Caption
                Close #1
                
                
                
                
                
            Shell App.Path & "\Reports\PrintSale.Bat", vbHide
        
            Open App.Path & "\Reports\Report.txt" For Output As #1
                Print #1, " "
            Close #1
                
                picmsg.SetFocus
                Timer1.Enabled = True
                txt(3) = ""
            End If
        End If
        
End Select
End Sub

Private Sub txtsearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        piclookup.Visible = False
        txt(0).Enabled = True
        txt(0).SetFocus
        lstsale.Enabled = True
        Toolbar1.Enabled = True
    End If
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
On Error Resume Next
'Procedur to search strings
'***********************************************************************
Dim inti As ListItem
If KeyAscii = 13 Then
    Set ac = New ADODB.Connection
    Set ar = New ADODB.Recordset
    
    Call dbconek
    ac.Open strConek
    
    lvsearch.ListItems.Clear
    With ar
        criteria = "Select *From itemmaster"
        .Open criteria, strConek, adOpenStatic, adLockOptimistic
            .MoveFirst
            Do While Not .EOF
                If Mid(!longdesc, 1, Len(txtsearch)) = txtsearch Then
                    Set intitem = lvsearch.ListItems.Add(, , !longdesc)
                    intitem.SubItems(1) = !upc
                End If
                .MoveNext
                lvsearch.SetFocus
                
            Loop
        .Close
    End With
End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
