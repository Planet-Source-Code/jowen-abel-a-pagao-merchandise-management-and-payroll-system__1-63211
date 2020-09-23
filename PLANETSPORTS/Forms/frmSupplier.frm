VERSION 5.00
Begin VB.Form frmMMSSupplier 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add and Update Supplier"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdlookup 
      Caption         =   "&Look Up"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&E&xit"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   4320
      Width           =   1095
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
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   9
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
      Height          =   285
      Index           =   3
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   8
      Top             =   2640
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
      Index           =   2
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2280
      Width           =   3975
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
      MaxLength       =   50
      TabIndex        =   6
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
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblins 
      BackStyle       =   0  'Transparent
      Caption         =   "For new supplier click &New. To edit a supplier click &Edit."
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User's Information"
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
      TabIndex        =   15
      Top             =   3360
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Supplier Information"
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
      TabIndex        =   14
      Top             =   1200
      Width           =   7095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7680
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Id:"
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
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1560
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
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. Number:"
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
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person:"
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
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "frmSupplier.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmMMSSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEdit_Click()
    Call txt_unlock
    cmdlookup.Enabled = True
    cmdsave.Enabled = True
    cmdsave.Caption = "&Update"
    cmdedit.Enabled = False
    cmdnew.Enabled = False
    Command2.Enabled = True
    txt(0).SetFocus
    
    lblins.Caption = "Enter the supplier ID then press enter or you can search for the supplier by name by clicking &Loop Up"
End Sub

Private Sub cmdlookup_Click()
    supplierflag = 1
    txt(0).SetFocus
    LoadForm SupplierSearch
End Sub

Private Sub cmdnew_Click()
    Call new_supplierid
    Call txt_unlock
    txt(0).Locked = True
    cmdnew.Enabled = False
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    Command2.Enabled = True
    txt(1).SetFocus
    
    lblins.Caption = "Fill up all the fields then press save."
End Sub

Private Sub cmdsave_Click()
Select Case cmdsave.Caption
    Case "&Save"
        Call dbconek
        With ar
            .Open "Select *From supplier", strConek, adOpenStatic, adLockOptimistic
                .AddNew
                    !supplierid = txt(0)
                    !supplierdesc = txt(1)
                    !contactperson = txt(2)
                    !address = txt(3)
                    !contactnumber = txt(4)
                .Update
                MsgBox "Supplier successfully added.", vbInformation, "New supplier added"
                Call Command2_Click
            .Close
        End With
    Case "&Update"
        Call dbconek
        With ar
            criteria = "Select *From supplier Where supplierid='" & txt(0) & "'"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    !supplierid = txt(0)
                    !supplierdesc = txt(1)
                    !contactperson = txt(2)
                    !address = txt(3)
                    !contactnumber = txt(4)
                .Update
                MsgBox "Supplier successfully updated.", vbInformation, "New supplier added"
                Call Command2_Click
            .Close
        End With
End Select
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
    LoadForm frmMMSSupplier
End Sub

Private Sub Form_Load()
    Me.Icon = frmNewEmployee.SmallImages.ListImages(2).Picture
End Sub

Private Sub new_supplierid()
On Error Resume Next
        Call dbconek
        
        
        Dim lastid As String
        With ar
            criteria = "Select *From supplier"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
            .MoveLast
            
            lastid = Mid(!supplierid, 5, Len(!supplierid))
            lastid = Val(lastid) + 1
            
            If Len(lastid) = 1 Then
                txt(0) = "SUP-" & "00" & lastid
            ElseIf Len(lastid) = 2 Then
                txt(0) = "SUP-" & "0" & lastid
            ElseIf Len(lastid) = 3 Then
                txt(0) = "SUP-" & lastid
            End If
            
            .Close
        End With
End Sub

Private Sub txt_unlock()
Dim i As Integer

    For i = 0 To 4
        txt(i).Locked = False
    Next
End Sub

Private Sub clear_txt()
Dim i As Integer

    For i = 0 To 4
        txt(i) = ""
    Next
End Sub
Private Sub txt_GotFocus(Index As Integer)
    SendKeys "{end}+{home}"
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0
        If KeyAscii = 13 Then
            Call dbconek
            
            With ar
                criteria = "Select *From supplier Where supplierid='" & txt(0) & "'"
                .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    If .RecordCount = 1 Then
                        txt(1) = !supplierdesc
                        txt(2) = !contactperson
                        txt(3) = !address
                        txt(4) = !contactnumber
                    Else
                        MsgBox "Supplier does not exist", vbInformation, "Search Error"
                        Exit Sub
                    End If
                .Close
            End With
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 1
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 2
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 3
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 4
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Select
End Sub
