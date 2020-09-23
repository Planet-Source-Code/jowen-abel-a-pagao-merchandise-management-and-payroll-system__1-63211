VERSION 5.00
Begin VB.Form frmMMSCategory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Category Maintenance"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmCategory.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Look Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmd 
      Caption         =   "F9 - &E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6120
      Picture         =   "frmCategory.frx":000C
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "F8 - &Delete"
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
      Height          =   375
      Index           =   2
      Left            =   4920
      Picture         =   "frmCategory.frx":09A6
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   5655
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
      Height          =   375
      Index           =   1
      Left            =   2520
      Picture         =   "frmCategory.frx":1340
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
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
      Height          =   375
      Index           =   0
      Left            =   3720
      Picture         =   "frmCategory.frx":1CDA
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
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
      Height          =   375
      Index           =   5
      Left            =   1320
      Picture         =   "frmCategory.frx":2674
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7440
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Press <ESC> to Cancel"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category Id:"
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
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "frmCategory.frx":300E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmMMSCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
        
                With ar
                    criteria = "Select *From category"
                    .Open criteria, strConek, adOpenStatic, adLockOptimistic
                        If Len(txt(1)) = 0 Then
                            MsgBox "Please enter category description", vbCritical, "Error Category"
                            Exit Sub
                        Else
                            .AddNew
                                !categoryid = UCase(txt(0))
                                !categorydesc = UCase(txt(1))
                            .Update
                            MsgBox "Category was successfully added", vbInformation, "Category Saved"
                            Unload Me
                            Load frmMMSCategory
                            frmMMSCategory.Show vbModal
                        End If
                    .Close
                End With
            Case "F4 - &Update"
                Set ac = New ADODB.Connection
                Set ar = New ADODB.Recordset
        
                Call dbconek
                ac.Open strConek
        
                With ar
                    criteria = "Select *From category Where categoryid='" & txt(0) & "'"
                    .Open criteria, strConek, adOpenStatic, adLockOptimistic
                        !categoryid = UCase(txt(0))
                        !categorydesc = UCase(txt(1))
                        .Update
                        MsgBox "Category was successfully updated", vbInformation, "Category Saved"
                        Unload Me
                        Load frmMMSCategory
                        frmMMSCategory.Show vbModal
                    .Close
                End With
        End Select
    Case 1
        Select Case cmd(1).Caption
            Case "F5 - &Edit"
                cmd(5).Enabled = False
                txt(0).SetFocus
                txt(0).Locked = False
                cmd(1).Caption = "F5 - &Search"
                cmd(0).Caption = "F4 - &Update"
            Case "F5 - &Search"
                                
                Set ac = New ADODB.Connection
                Set ar = New ADODB.Recordset
    
                Call dbconek
                ac.Open strConek
                        
                With ar
                    criteria = "Select *From category Where categoryid='" & txt(0) & "'"
                    .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    If .RecordCount >= 1 Then
                        cmd(0).Enabled = True
                        cmd(2).Enabled = True
                        txt(1) = !categorydesc
                    Else
                        MsgBox "Category does not exist", vbInformation, "Category Error"
                        txt(0) = ""
                        txt(0).SetFocus
                        Exit Sub
                    End If
                    .Close
                End With
        End Select
    Case 2
        Set ac = New ADODB.Connection
        Set ar = New ADODB.Recordset
        
        Call dbconek
        ac.Open strConek
        
        Dim dans
        
        dans = MsgBox("Are you sure you want to delete " & txt(1) & " category?", vbQuestion + vbYesNo, "Delete Alert")
        If dans = vbYes Then
            With ar
                criteria = "Select *From category Where categoryid='" & UCase(txt(0)) & "'"
                .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    .Delete
                    MsgBox "Category " & txt(1) & " was successfully deleted", vbInformation, "Delete Success"
                    Unload Me
                    Load frmMMSCategory
                    frmMMSCategory.Show vbModal
                .Close
            End With
        Else
            Cancel = 1
            Exit Sub
        End If
    Case 3
        Unload Me
    Case 5
        Set ac = New ADODB.Connection
        Set ar = New ADODB.Recordset
    
        Call dbconek
        ac.Open strConek
        
        Dim lastid As String
        With ar
            criteria = "Select *From category"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
            .MoveLast
            
            lastid = Mid(!categoryid, 5, Len(!categoryid))
            lastid = Val(lastid) + 1
            
            If Len(lastid) = 1 Then
                txt(0) = "CAT-" & "00" & lastid
            ElseIf Len(lastid) = 2 Then
                txt(0) = "CAT-" & "0" & lastid
            ElseIf Len(lastid) = 3 Then
                txt(0) = "CAT-" & lastid
            End If
            txt(1).SetFocus
            .Close
        End With
        cmd(0).Enabled = True
        cmd(1).Enabled = False
    
End Select
End Sub

Private Sub Form_Load()
    Me.Icon = frmNewEmployee.SmallImages.ListImages(3).Picture
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
    Case vbKeyEscape
        Unload Me
        Load frmMMSCategory
        frmMMSCategory.Show vbModal
End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 1
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Select
End Sub
