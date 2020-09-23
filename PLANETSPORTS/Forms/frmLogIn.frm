VERSION 5.00
Begin VB.Form frmxLogIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogIn.frx":0000
   ScaleHeight     =   3630
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "F3 - &Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4200
      Picture         =   "frmLogIn.frx":2B2B
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "F6 - &Log In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2640
      Picture         =   "frmLogIn.frx":34C5
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txt 
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
      Index           =   1
      Left            =   2640
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox txt 
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
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   0
      Picture         =   "frmLogIn.frx":3E5F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6180
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   1365
      Left            =   360
      Picture         =   "frmLogIn.frx":65EA
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1725
   End
End
Attribute VB_Name = "frmxLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim passattemp As Integer
Private Sub cmd_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
    'Check if the username and the password is correct
        If txt(0) = "" Then
            MsgBox "Please enter your Username", vbInformation, "Log In Error"
            txt(0).SetFocus
            Exit Sub
        End If
                
        Set ac = New ADODB.Connection
        Set ar = New ADODB.Recordset
        
        Call dbconek
        ac.Open strConek
        
        With ar
            criteria = "Select *From systemuser Where username='" & txt(0) & "'"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
            If .RecordCount >= 1 Then
                If !Password = txt(1) Then
                    currentemmsuser = !Name
                    CurrentPosition = !jobposition
                    currentjoblevel = !joblevel
                    If CurrentPosition = "CASHIER" Then
                        Unload Me
                        LoadForm posterminal
                    Else
                        Unload Me
                        Load MAIN
                        MAIN.Show
                    End If
                Else
                    passattemp = passattemp + 1
                    If passattemp = 3 Then
                        MsgBox "You are not an authorized user", vbInformation + vbCritical, "Log In Error"
                        Unload Me
                    Else
                        MsgBox "Password incorrect. Please check the CAPS LOCK" & vbCrLf & "                        Attempt left " & 3 - passattemp & "", vbInformation, "Log In Error"
                        txt(1) = ""
                        txt(1).SetFocus
                    End If
                End If
            Else
                MsgBox "This user does not exist", vbCritical, "Log In Error"
                txt(0) = ""
                txt(0).SetFocus
            End If
            .Close
        End With
        
        Call dbconek
        
        With ar
            criteria = "Select *From userhistory"
            .Open criteria, strConek, 3, 3
                .AddNew
                    !nameofuser = currentemmsuser
                    !timein = Time
                    !Date = Date
                .Update
            .Close
        End With
    Case 1
        Unload Me
End Select
End Sub

Private Sub Form_Load()
    Me.Icon = frmNewEmployee.SmallImages.ListImages(23).Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call dbclose
End Sub

'Event when the return key or the enter key is depress'
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Select Case Index
    Case 0
        If KeyAscii = 13 Then
            If txt(0) = "" Then
                MsgBox "Please enter your Username", vbInformation, "Log In Error"
                txt(0).SetFocus
                Exit Sub
            
            Else
                txt(1).SetFocus
                Exit Sub
            End If
        End If
    Case 1
        If KeyAscii = 13 Then
            If txt(1) = "" Then
                MsgBox "Please enter your Password", vbInformation, "Log In Error"
                txt(1).SetFocus
                Exit Sub
            
            Else
                cmd(0).SetFocus
                Exit Sub
            End If
        End If
End Select
End Sub
