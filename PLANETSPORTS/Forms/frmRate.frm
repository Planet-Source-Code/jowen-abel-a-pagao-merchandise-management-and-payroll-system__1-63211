VERSION 5.00
Begin VB.Form frmRate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate and Position Maintenance"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   Icon            =   "frmRate.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   9420
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
      Height          =   270
      Index           =   3
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   3
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
      Index           =   0
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   2400
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
      Index           =   4
      Left            =   6720
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2400
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
      Left            =   6720
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   2760
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
      ItemData        =   "frmRate.frx":000C
      Left            =   2160
      List            =   "frmRate.frx":000E
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
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
      Height          =   285
      Index           =   1
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "F9 - &Cancel"
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
      Index           =   3
      Left            =   7800
      Picture         =   "frmRate.frx":0010
      TabIndex        =   12
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "F8 - &Delete"
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
      Index           =   2
      Left            =   6360
      Picture         =   "frmRate.frx":09AA
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "F5 - &Edit"
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
      Left            =   3480
      Picture         =   "frmRate.frx":1344
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "F4 - &Save"
      Enabled         =   0   'False
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
      Left            =   4920
      Picture         =   "frmRate.frx":1CDE
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "F6 - &New"
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
      Index           =   5
      Left            =   2040
      Picture         =   "frmRate.frx":2678
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
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
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9480
      Y1              =   3240
      Y2              =   3240
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
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   4335
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
      Left            =   4800
      TabIndex        =   18
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holidays:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   3
      Left            =   4800
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
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
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   16
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holidays:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   1815
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
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   2535
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
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
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
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   8
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "frmRate.frx":3012
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbox_Click()
On Error Resume Next
                With ar
                    criteria = "Select *From rateposition Where jobposition='" & cbox & "'"
                    .Open criteria, strConek, adOpenStatic, adLockOptimistic
                        
                        txt(0) = Format(!rateperhour, "###,###,##0.00")
                        txt(3) = Format(!specialrate, "###,###,##0.00")
                        txt(4) = Format(!otregdays, "###,###,##0.00")
                        txt(5) = Format(!otspecial, "###,###,##0.00")
                        txt(1) = !joblevel
                    .Close
                End With
End Sub

Private Sub cbox_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF4
        Call cmd_Click(0)
    Case vbKeyF5
        Call cmd_Click(1)
    Case vbKeyF6
        Call cmd_Click(5)
    Case vbKeyF8
        Call cmd_Click(2)
    Case vbKeyF9
        Call cmd_Click(3)
End Select
End Sub

Private Sub cmd_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0
            Select Case cmd(0).Caption
            Case "F4 - &Save"
                Set ac = New ADODB.Connection
                Set ar = New ADODB.Recordset
        
                Call dbconek
                ac.Open strConek
                If txt(2) = "" Then Exit Sub
                With ar
                    criteria = "Select *From rateposition"
                    .Open criteria, strConek, adOpenStatic, adLockOptimistic
                        If Len(txt(2)) = 0 Then
                            MsgBox "Please enter position description", vbCritical, "Error Category"
                            Exit Sub
                        Else
                            .AddNew
                                !jobposition = txt(2)
                                !rateperhour = txt(0)
                                !joblevel = txt(1)
                                !specialrate = txt(3)
                                !otregdays = txt(4)
                                !otspecial = txt(5)
                            .Update
                            MsgBox "Category was successfully added", vbInformation, "Category Saved"
                            Unload Me
                            Load frmRate
                            frmRate.Show vbModal
                        End If
                    .Close
                End With
            Case "F4 - &Update"
                Set ac = New ADODB.Connection
                Set ar = New ADODB.Recordset
        
                Call dbconek
                ac.Open strConek
                
                'If Len(txt(2)) = 0 Then Exit Sub
                With ar
                    criteria = "Select *From rateposition Where jobposition='" & txt(2).Text & "'"
                    .Open criteria, strConek, adOpenStatic, adLockOptimistic
                        
                        If txt(0).Text = "" Or txt(1) = " " Then
                            Exit Sub
                        End If
                            !rateperhour = txt(0)
                            !joblevel = txt(1)
                            !specialrate = txt(3)
                            !otregdays = txt(4)
                            !otspecial = txt(5)
                        .Update
                        MsgBox "Category was successfully updated", vbInformation, "Category Saved"
                        Unload Me
                        Load frmRate
                        frmRate.Show vbModal
                    .Close
                End With
        End Select
        Case 1
            txt(0).Locked = False
            txt(1).Locked = False
            txt(0).SetFocus
            SendKeys "{home}+{end}"
            cmd(0).Enabled = True
            cmd(0).Caption = "F4 - &Update"
            cmd(5).Enabled = False
            
        Case 2
            Set ac = New ADODB.Connection
                Set ar = New ADODB.Recordset
        
                Call dbconek
                ac.Open strConek
                
                If Len(txt(2)) = 0 Then Exit Sub
                With ar
                    criteria = "Select *From rateposition Where jobposition='" & txt(2).Text & "'"
                    .Open criteria, strConek, adOpenStatic, adLockOptimistic
                        
                        If txt(0).Text = "" Or txt(1) = " " Then
                            Exit Sub
                        End If
                        !rateperhour = txt(0)
                        !joblevel = txt(1)
                        .Delete
                        MsgBox "Category was successfully updated", vbInformation, "Category Saved"
                        Unload Me
                        Load frmRate
                        frmRate.Show vbModal
                    .Close
                End With
        Case 3
            Unload Me
        Case 5
            cbox.Visible = False
            txt(2).Visible = True
            txt(2).SetFocus
            cmd(1).Enabled = False
            cmd(2).Enabled = False
            txt(0).Locked = False
            txt(1).Locked = False
            cmd(0).Enabled = True
            cmd(0).Caption = "F4 - &Save"
    End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Me.Icon = frmNewEmployee.SmallImages.ListImages(5).Picture
    
    Set ac = New ADODB.Connection
    Set ar = New ADODB.Recordset
    

    Call dbconek
    ac.Open strConek
    
    With ar
        criteria = "Select *From rateposition"
        .Open criteria, strConek, adOpenStatic, adLockOptimistic
        .MoveFirst
        Do While Not .EOF
            cbox.AddItem !jobposition
            .MoveNext
        Loop
        
        .Close
    End With
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SendKeys "{end}+{home}"
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF4
        Call cmd_Click(0)
    Case vbKeyF5
        Call cmd_Click(1)
    Case vbKeyF6
        Call cmd_Click(5)
    Case vbKeyF8
        Call cmd_Click(2)
    Case vbKeyF9
        Call cmd_Click(3)
End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    Case 1
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    Case 2
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 3
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    Case 4
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    Case 5
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
End Select
End Sub
