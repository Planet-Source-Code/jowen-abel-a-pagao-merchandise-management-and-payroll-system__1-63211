VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   855
      HideSelection   =   0   'False
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmAbout.frx":000C
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.00 Copyright 2005"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning :"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   6
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Sports LCC Branch NagaCity"
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   4
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "This software is license to:"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "with Point of Sales Terminal"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":00F9
      Height          =   975
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee and Merchandise Management System"
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   3180
      Left            =   120
      Picture         =   "frmAbout.frx":01B5
      Top             =   1320
      Width           =   2805
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "frmAbout.frx":1D507
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmShortcuts.SmallImages.ListImages(5).Picture
End Sub

