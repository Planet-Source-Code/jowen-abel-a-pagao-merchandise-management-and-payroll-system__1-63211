VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   4
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   2640
      Width           =   1455
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
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   3000
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   2
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1560
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2280
      Width           =   4695
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
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Month/s"
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
      Left            =   3600
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Deduction"
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
      Left            =   360
      TabIndex        =   15
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Term :"
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
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Amount :"
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
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
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
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Number :"
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
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Loanee Information and Mode of Payment"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbox_Click(Index As Integer)
    txt(4).Text = Format(CCur(txt(3) / cbox(0)), "###,###,##0.00")
End Sub

Private Sub Command1_Click()
On Error Resume Next
    Command4.Enabled = True
    Command2.Enabled = True
    Command1.Enabled = False
    
    Call dbconek
    
    ar.Open "Select *From loantable", strConek, adOpenStatic, adLockOptimistic
        ar.MoveLast
            txt(2) = ar!loannumber + 1
            txt(0).SetFocus
    ar.Close
End Sub

Private Sub Command2_Click()
    If txt(3) = "0.00" Or txt(4) = "0.00" Then
        txt(3).SetFocus
        Exit Sub
    End If
    
    Call dbconek
    
    With ar
        .Open "Select *From loantable", strConek, adOpenStatic, adLockOptimistic
            .AddNew
                !loannumber = txt(2)
                !loanamount = txt(3)
                !deductionpermonth = txt(4)
                !employeeid = txt(0)
                !employeename = txt(1)
                !loanterm = cbox(0)
                !paid = "NO"
            .Update
            MsgBox "Loan Successfull...", vbInformation
            
        .Close
        Call Command4_Click
    End With
        
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
    LoadForm Form1
End Sub

Private Sub Form_Load()
    cbox(0).AddItem "1"
    cbox(0).AddItem "2"
    cbox(0).AddItem "3"
    cbox(0).AddItem "4"
    cbox(0).AddItem "5"
    cbox(0).AddItem "6"
    cbox(0).AddItem "7"
    cbox(0).AddItem "8"
    cbox(0).AddItem "9"
    cbox(0).AddItem "10"
    cbox(0).AddItem "11"
    cbox(0).AddItem "12"
End Sub

Private Sub txt_Click(Index As Integer)
    SendKeys "{end}+{home}"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SendKeys "{end}+{home}"
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Select Case Index
    Case 0
        If KeyAscii = 13 Then
            Call dbconek
            ar.Open "Select *From employeemaster Where employeeid='" & txt(0) & "'", strConek, adOpenStatic, adLockOptimistic
                If ar.RecordCount = 1 Then
                    txt(1) = ar!firstname & " " & ar!middlename & " " & ar!lastname
                    txt(3).SetFocus
                Else
                    txt(0).SetFocus
                    Exit Sub
                End If
            ar.Close
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 3
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
End Select
End Sub
