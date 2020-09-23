VERSION 5.00
Begin VB.Form frmUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System User Maintenance"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   21
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   4560
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
      Index           =   6
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   1
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   1920
      PasswordChar    =   "l"
      TabIndex        =   3
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1920
      PasswordChar    =   "l"
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
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
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1455
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
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
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
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1920
      Width           =   5535
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
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Security Information"
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
      Left            =   240
      TabIndex        =   22
      Top             =   3000
      Width           =   7215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(16 Chars. Max.)"
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
      Left            =   4920
      TabIndex        =   16
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(16 Chars. Max.)"
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
      Left            =   4920
      TabIndex        =   15
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
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
      Left            =   240
      TabIndex        =   13
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Job Level:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Position:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Information"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "frmUser.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdedit_Click()
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    cmdsave.Caption = "&Update"
    Command1.Enabled = True
    cmdnew.Enabled = False
    txt(0).SetFocus
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdnew_Click()
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    Command1.Enabled = True
    cmdnew.Enabled = False
    txt(0).SetFocus
    
End Sub

Private Sub cmdsave_Click()
On Error Resume Next
Select Case cmdsave.Caption
    Case "&Save"
        If txt(4) = "" Or txt(6) = "" Then
            txt(4).SetFocus
            Exit Sub
        End If
        If txt(4) <> txt(5) Then
            MsgBox "Please Confirm Your Pawword", vbCritical, "Password Confirmation Error"
            txt(5).SetFocus
            Exit Sub
        End If
        Call dbconek
        With ar
            .Open "Select *From systemuser", strConek, adOpenStatic, adLockOptimistic
                .AddNew
                    !employeeid = txt(0)
                    !Name = txt(1)
                    !jobposition = txt(2)
                    !joblevel = txt(3)
                    !UserName = txt(6)
                    !Password = txt(4)
                .Update
                MsgBox "User successfully added", vbInformation, "Save Success"
                Call Command1_Click
            .Close
        End With
    Case "&Update"
        If txt(4) = "" Or txt(6) = "" Then
            txt(4).SetFocus
            Exit Sub
        End If
        If txt(4) <> txt(5) Then
            MsgBox "Please Confirm Your Pawword", vbCritical, "Password Confirmation Error"
            txt(5).SetFocus
            Exit Sub
        End If
        Call dbconek
        With ar
            .Open "Select *From systemuser Where employeeid='" & txt(0) & "'", strConek, adOpenStatic, adLockOptimistic
                    !Name = txt(1)
                    !jobposition = txt(2)
                    !joblevel = txt(3)
                    !UserName = txt(6)
                    !Password = txt(4)
                .Update
                MsgBox "User successfully updated", vbInformation, "Update Success"
                Call Command1_Click
            .Close
        End With
End Select
End Sub

Private Sub Command1_Click()
    Unload Me
    LoadForm frmUser
End Sub

Private Sub Form_Load()
    Me.Icon = frmShortcuts.ImageList1.ListImages(10).Picture
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Select Case Index
    Case 0
        
        If KeyAscii = 13 Then
        Call dbconek
        
        With ar
            .Open "Select *From employeemaster Where employeeid='" & txt(0) & "'", strConek, adOpenStatic, adLockOptimistic
                If .RecordCount >= 1 Then
                    txt(1) = !firstname & " " & !middlename & " " & !lastname
                    txt(2) = !jobposition
                Else
                    txt(0).SetFocus
                    SendKeys "{end}+{home}"
                    Exit Sub
                End If
            .Close
        End With
        
        Call dbconek
        
        With ar
            .Open "Select *From rateposition Where jobposition='" & txt(2) & "'", strConek, adOpenStatic, adLockOptimistic
                If .RecordCount >= 1 Then
                    txt(3) = !joblevel
                End If
                txt(6).SetFocus
            .Close
        End With
        End If
        
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Select
End Sub

