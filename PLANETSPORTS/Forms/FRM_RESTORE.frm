VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRM_RESTORE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore Database ...."
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   Icon            =   "FRM_RESTORE.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit "
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton CMD_RESTORE 
      Caption         =   "&Restore "
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   5640
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   4800
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   5160
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Shape Shape1 
      Height          =   6255
      Left            =   0
      Top             =   0
      Width           =   7725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Restore Data Base"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   4680
      TabIndex        =   9
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   3465
      Left            =   120
      Picture         =   "FRM_RESTORE.frx":0442
      Top             =   1680
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Backed Up File and Click on  Restore Button"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   4320
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selected Back Up File Path"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "FRM_RESTORE.frx":4902
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7740
   End
End
Attribute VB_Name = "FRM_RESTORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1

Private Sub CMD_RESTORE_Click()
  If Len(Label1(1).Caption) = 0 Then
    MsgBox "Select Backup File Path and then Click Restore Button...", vbInformation, "Select Path ..."
    Exit Sub
  End If
  
  Dim OldTimer As Single
  ProgressBar1.Visible = True
  Label3.Visible = True
  
  OldTimer = Time
  Call Huffman.DecodeFile(Label1(1).Caption, App.Path & "\dbase_backup_bk.mdb")
  
  ProgressBar1.Value = 0
  
  Unload Me
  Exit Sub
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
Label1(1).Caption = ""
On Error GoTo A1:
    File1.Path = Dir1.Path
    Exit Sub
A1:
    MsgBox "Folder Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"
End Sub

Private Sub Drive1_Change()
Label1(1).Caption = ""
On Error GoTo A1:
    Dir1.Path = Drive1.Drive
    Exit Sub
A1:
    MsgBox "Drive Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"

End Sub

Private Sub File1_Click()
    Label1(1).Caption = File1.Path & "\" & File1.Filename
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If ProgressBar1.Value = 0 Or ProgressBar1.Value Then
        Unload Me
    End If
End If

End Sub

Private Sub Form_Load()
KeyPreview = True

Me.Top = 0
Me.Left = 0
Label1(1).Caption = ""
Label3.Visible = False
ProgressBar1.Visible = False
Set Huffman = New clsHuffman
Drive1.Drive = "c:"
End Sub

Private Sub Huffman_Progress(Procent As Integer)
  Label3.Caption = "Uncompressing Database"
  ProgressBar1.Value = Procent
  If ProgressBar1.Value = 100 Then
    Label3.Caption = "Restoring Uncompressed File ..."
    End If
  DoEvents
End Sub

