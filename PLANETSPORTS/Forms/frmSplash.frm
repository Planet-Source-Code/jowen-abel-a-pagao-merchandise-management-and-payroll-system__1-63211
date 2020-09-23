VERSION 5.00
Begin VB.Form frmxSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4410
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4410
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   600
      Top             =   2880
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   4410
      Left            =   0
      Top             =   0
      Width           =   8205
   End
   Begin VB.Image Image1 
      Height          =   4410
      Left            =   0
      Picture         =   "frmSplash.frx":76056
      Top             =   0
      Width           =   8220
   End
End
Attribute VB_Name = "frmxSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim RetValue
    'RetValue = ChangeRes(1024, 768, 32)
End Sub

Private Sub Timer1_Timer()
    
    Unload Me
    Load frmxDateTime
    frmxDateTime.Show vbModal
    
End Sub
