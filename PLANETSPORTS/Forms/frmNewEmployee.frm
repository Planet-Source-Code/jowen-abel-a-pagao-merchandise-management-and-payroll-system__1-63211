VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmNewEmployee 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add and Update Employee "
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13605
   Icon            =   "frmNewEmployee.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   15
      Left            =   6480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   53
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   14
      Left            =   10920
      MaxLength       =   15
      TabIndex        =   49
      Top             =   4080
      Width           =   2415
   End
   Begin VB.ComboBox cboxph 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10920
      TabIndex        =   48
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   44
      Text            =   "0.00"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   12
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   43
      Text            =   "0.00"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   40
      Text            =   "0.00"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   9
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   39
      Text            =   "0.00"
      Top             =   1920
      Width           =   2415
   End
   Begin VB.ComboBox cboxtin 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   12
      Top             =   4440
      Width           =   2415
   End
   Begin VB.ComboBox cboxsss 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   37
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdlookup 
      Caption         =   "&Look Up"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   34
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
      Height          =   375
      Left            =   8160
      TabIndex        =   33
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   9480
      TabIndex        =   32
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   12120
      TabIndex        =   31
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10800
      TabIndex        =   29
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   11
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   11
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   10
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   5
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox cbox 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":000C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":0CE6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":1B38
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":298A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":3264
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":3B3E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":4418
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":4DE2
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":56BC
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":59D6
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":62B0
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":6B8A
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":7464
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":777E
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":8058
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":8932
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":920C
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":9AE6
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":A3C0
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":AC9A
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":B574
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":BE4E
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":C728
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":D002
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":D8DC
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":E1B6
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":EA90
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":F36A
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":FC44
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":1051E
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":10DD4
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":116AE
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":11B00
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":11F52
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewEmployee.frx":14704
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4800
      TabIndex        =   54
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "P. Health Number :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   9240
      TabIndex        =   52
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "P.H. Bracket :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   9240
      TabIndex        =   51
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Phil. Health Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   9240
      TabIndex        =   50
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hourly Rates for Overtime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9240
      TabIndex        =   47
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Days:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   9240
      TabIndex        =   46
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holidays :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   9240
      TabIndex        =   45
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Days:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   9240
      TabIndex        =   42
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holidays :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   9240
      TabIndex        =   41
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hourly Rates for Days Work"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9240
      TabIndex        =   38
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lblins 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNewEmployee.frx":15986
      Height          =   495
      Left            =   360
      TabIndex        =   36
      Top             =   5160
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User's Instruction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   35
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   4800
      TabIndex        =   30
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   1  'Blackness
      X1              =   0
      X2              =   13560
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "With Holding Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   28
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TAX Bracket :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   4800
      TabIndex        =   27
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TIN Number :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   4800
      TabIndex        =   26
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Social Security Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   25
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Bracket :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   360
      TabIndex        =   24
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Number :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   360
      TabIndex        =   23
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Position :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   4800
      TabIndex        =   22
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Start of Contract :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4800
      TabIndex        =   21
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "End of Contract :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4800
      TabIndex        =   20
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Employement Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   19
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "frmNewEmployee.frx":15A16
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13740
   End
End
Attribute VB_Name = "frmNewEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbox_Click()
On Error Resume Next
    Call dbconek
    
    ar.Open "Select *From rateposition Where jobposition='" & cbox.Text & "'", strConek, adOpenStatic, adLockOptimistic
        txt(8) = Format(CCur(ar!rateperhour), "###,##0.00")
        txt(9) = Format(CCur(ar!specialrate), "###,##0.00")
        txt(12) = Format(CCur(ar!otregdays), "###,##0.00")
        txt(13) = Format(CCur(ar!otspecial), "###,##0.00")
    ar.Close
End Sub

Private Sub cbox_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDown Or KeyAscii = vbKeyUp Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub cboxph_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDown Or KeyAscii = vbKeyUp Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub cboxsss_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDown Or KeyAscii = vbKeyUp Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub cboxtin_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDown Or KeyAscii = vbKeyUp Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub cmdedit_Click()
    cmdsave.Enabled = True
    cmdsave.Caption = "&Update"
    Call txt_unlock
    Command1.Enabled = True
    cmdedit.Enabled = False
    cmdlookup.Enabled = True
    txt(5).SetFocus
    lblins.Caption = "Type the employee ID and then press <Enter Key>. If you have forgotten the employeeID you can click &LookUp button to search for the employee you want to Update. "
    
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdlookup_Click()
    employeesearh = 1
    txt(5).SetFocus
    LoadForm employeesearch
End Sub

Private Sub cmdnew_Click()
    Call new_employeeid
    Call txt_unlock
    
    cmdsave.Enabled = True
    cmdedit.Enabled = False
    cmdlookup.Enabled = False
    Command1.Enabled = True
    cmdnew.Enabled = False
    txt(5).Locked = True
    txt(0).SetFocus
    
    lblins.Caption = "After you have finished typing on the fields you can click &Save button to save the new employee. All fields are required to have a value."
    
    
End Sub

Private Sub cmdsave_Click()
'On Error Resume Next
    Select Case cmdsave.Caption
        Case "&Save"
            For i = 0 To 11
                If txt(i) = "" Then
                    MsgBox "All fields are required", vbInformation, "Save Error"
                    txt(i).SetFocus
                    Exit Sub
                End If
            Next
            With ar
                criteria = "Select * From employeemaster"
                .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    .AddNew
                    !lastname = txt(0)
                    !firstname = txt(1)
                    !middlename = txt(2)
                    !dateofbirth = txt(3)
                    !address = txt(4)
                    !employeeid = txt(5)
                    !jobposition = cbox
                    !startofcontract = txt(6)
                    !endofcontract = txt(7)
                    !rateperhour = txt(8)
                    !specialrate = txt(9)
                    !sssnumber = txt(10)
                    !tinnumber = txt(11)
                    !sssbracket = cboxsss
                    !taxbracket = cboxtin
                    !phnumber = txt(14)
                    !phbracket = cboxph
                    !Comments = txt(15)
                    !otregdays = txt(12)
                    !otspecial = txt(13)
                    .Update
                    
                    A = MsgBox("Employee have been successfully added. Do you want to add another?", vbYesNo + vbQuestion, "Add Success")
                    If A = vbYes Then
                        Call clear_txt
                        Call cmdnew_Click
                    Else
                        Call Command1_Click
                    End If
                .Close
            End With
        Case "&Update"
            For i = 0 To 11
                If txt(i) = "" Then
                    MsgBox "All fields are required", vbInformation, "Save Error"
                    txt(i).SetFocus
                    Exit Sub
                End If
            Next
            With ar
                criteria = "Select * From employeemaster Where employeeid='" & txt(5) & "'"
                .Open criteria, ctrconek, adOpenStatic, adLockOptimistic
                    !lastname = txt(0)
                    !firstname = txt(1)
                    !middlename = txt(2)
                    !dateofbirth = txt(3)
                    !address = txt(4)
                    !employeeid = txt(5)
                    !jobposition = cbox
                    !startofcontract = txt(6)
                    !endofcontract = txt(7)
                    !rateperhour = txt(8)
                    !specialrate = txt(9)
                    !sssnumber = txt(10)
                    !tinnumber = txt(11)
                    !sssbracket = cboxsss
                    !taxbracket = cboxtin
                    !phnumber = txt(14)
                    !phbracket = cboxph
                    !Comments = txt(15)
                    !otregdays = txt(12)
                    !otspecial = txt(13)
                    .Update
                    b = MsgBox("Employee have been successfully updated. Do you want to update another?", vbYesNo + vbQuestion, "Add Success")
                    If b = vbYes Then
                        Call clear_txt
                        Call cmdedit_Click
                    Else
                        Call Command1_Click
                    End If
                    
                .Close
            End With
    End Select
End Sub

Private Sub Command1_Click()
    Unload Me
    LoadForm frmNewEmployee
    
End Sub

Private Sub Form_Load()
Dim mydate As Integer
Dim myyear As Double
    
    Call fillcbox1
    
    Me.Icon = Me.SmallImages.ListImages(19).Picture
    Image1.Width = Me.Width
    mydate = Format(Date, "mm") + 3
    myyear = Format(Date, "yyyy")
    If mydate > 12 Then
        txt(7) = Format(Date, mydate - 12 & "/dd/" & myyear + 1)
    Else
        txt(7) = Format(Date, mydate & "/dd/yyyy")
    End If
    txt(6) = Format(Date, "mm/dd/yyyy")
    
    Call fillcbox
End Sub

Private Sub fillcbox()
On Error Resume Next
    Call dbconek
        
    With ar
        criteria = "Select *From rateposition"
        .Open criteria, strConek, 3, 3
            .MoveFirst
            Do While Not .EOF
                cbox.AddItem !jobposition
                .MoveNext
            Loop
        .Close
    End With
End Sub

Private Sub new_employeeid()
On Error Resume Next
        Call dbconek
        
        
        Dim lastid As String
        With ar
            criteria = "Select *From employeemaster"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
            .MoveLast
            
            lastid = Mid(!employeeid, 5, Len(!employeeid))
            lastid = Val(lastid) + 1
            
            If Len(lastid) = 1 Then
                txt(5) = "EMP-" & "00" & lastid
            ElseIf Len(lastid) = 2 Then
                txt(5) = "EMP-" & "0" & lastid
            ElseIf Len(lastid) = 3 Then
                txt(5) = "EMP-" & lastid
            End If
            
            .Close
        End With
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SendKeys "{end}+{home}"
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 1
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 2
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 3
        
            Dim ss As String
            If Len(txt(3)) < 2 Then
                ss = "0123456789" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(3), ss)
            ElseIf Len(txt(3)) = 2 Then
                ss = "/" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(3), ss)
            ElseIf Len(txt(3)) = 3 Then
                ss = "0123456789" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(3), ss)
            ElseIf Len(txt(3)) = 5 Then
                ss = "/" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(3), ss)
            ElseIf Len(txt(3)) = 6 Then
                ss = "0123456789" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(3), ss)
            End If
    Case 4
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 5
        If KeyAscii = 13 Then
            Call dbconek
            
            With ar
                criteria = "Select * From employeemaster Where employeeid='" & txt(5).Text & "'"
                .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    If .RecordCount >= 1 Then
                    txt(0) = !lastname
                    txt(1) = !firstname
                    txt(2) = !middlename
                    txt(3) = !dateofbirth
                    txt(4) = !address
                    txt(5) = !employeeid
                    cbox = !jobposition
                    txt(6) = !startofcontract
                    txt(7) = !endofcontract
                    txt(8) = !rateperhour
                    txt(9) = !otrate
                    txt(10) = !sssnumber
                    txt(11) = !tinnumber
                    cboxsss = !sssbracket
                    cboxtin = !taxbracket
                    Else
                        MsgBox "Employee does not exist. Please try to search using part of employee's name.", vbInformation, "Search Error"
                        Exit Sub
                    End If
                .Close
            End With
        End If
    Case 6
            Dim sS1 As String
            If Len(txt(6)) < 2 Then
                sS1 = "0123456789" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(6), sS1)
            ElseIf Len(txt(6)) = 2 Then
                sS1 = "/" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(6), sS1)
            ElseIf Len(txt(6)) = 3 Then
                sS1 = "0123456789" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(6), sS1)
            ElseIf Len(txt(6)) = 5 Then
                sS1 = "/" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(6), sS1)
            ElseIf Len(txt(6)) = 6 Then
                sS1 = "0123456789" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(6), sS1)
            End If
    Case 7
            Dim sS2 As String
            If Len(txt(7)) < 2 Then
                sS2 = "0123456789" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(7), sS2)
            ElseIf Len(txt(7)) = 2 Then
                sS2 = "/" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(7), sS2)
            ElseIf Len(txt(7)) = 3 Then
                sS2 = "0123456789" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(7), sS2)
            ElseIf Len(txt(7)) = 5 Then
                sS2 = "/" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(7), sS2)
            ElseIf Len(txt(7)) = 6 Then
                sS2 = "0123456789" & Chr(vbKeySpace) & Chr(vbKeyBack)

                Call offDefine(KeyAscii, txt(7), sS2)
            End If
    Case 8
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    Case 9
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    Case 10
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 11
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 12
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    Case 13
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
End Select
End Sub

Private Sub txt_unlock()
Dim i As Integer

    For i = 0 To 11
        txt(i).Locked = False
    Next
End Sub

Private Sub clear_txt()
Dim i As Integer

    For i = 0 To 11
        txt(i) = ""
    Next
End Sub

Private Sub fillcbox1()
On Error Resume Next
    Call dbconek
    
    ar.Open "Select *From ssstable", strConek, adOpenStatic, adLockOptimistic
        ar.MoveFirst
        Do While Not ar.EOF
            cboxsss.AddItem ar!salarybracket
            ar.MoveNext
        Loop
    ar.Close
    
    Call dbconek
    
    ar.Open "Select *From phtable", strConek, adOpenStatic, adLockOptimistic
        ar.MoveFirst
        Do While Not ar.EOF
            cboxph.AddItem ar!salarybracket
            ar.MoveNext
        Loop
    ar.Close
    
    Call dbconek
    
    ar.Open "Select *From taxtable", strConek, adOpenStatic, adLockOptimistic
        ar.MoveFirst
        Do While Not ar.EOF
            cboxtin.AddItem ar!salarybracket
            ar.MoveNext
        Loop
    ar.Close
End Sub



