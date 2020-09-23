VERSION 5.00
Begin VB.Form frmMMSitem 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Item Maintenance"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   2145
   ClientWidth     =   8115
   Icon            =   "frmItem.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleMode       =   0  'User
   ScaleWidth      =   8111.379
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   27
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6720
      TabIndex        =   26
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   4080
      TabIndex        =   25
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdlookup 
      Caption         =   "&Look Up"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   5640
      Width           =   1215
   End
   Begin VB.ComboBox cboxinterest 
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
      Left            =   6120
      TabIndex        =   19
      Text            =   "0.01"
      Top             =   4200
      Width           =   1695
   End
   Begin VB.ComboBox cboxdiscount 
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
      Left            =   6120
      TabIndex        =   18
      Text            =   "0.01"
      Top             =   3840
      Width           =   1695
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
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   4200
      Width           =   1575
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
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   3840
      Width           =   1575
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
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   3000
      Width           =   2655
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
      TabIndex        =   7
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
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   5
      Top             =   2280
      Width           =   5775
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
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1920
      Width           =   5775
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
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lblcat 
      Caption         =   "lblcat"
      Height          =   255
      Left            =   4800
      TabIndex        =   29
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblsup 
      Caption         =   "Label3"
      Height          =   255
      Left            =   4800
      TabIndex        =   28
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8156.359
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label lblins 
      BackStyle       =   0  'Transparent
      Caption         =   "For new product/item click &New. To edit a product/item click &Edit."
      Height          =   495
      Left            =   360
      TabIndex        =   21
      Top             =   5040
      Width           =   7455
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
      Index           =   2
      Left            =   360
      TabIndex        =   20
      Top             =   4680
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Interest in Credit:"
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
      Left            =   4560
      TabIndex        =   17
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount:"
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
      Left            =   4560
      TabIndex        =   16
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price :"
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
      Left            =   360
      TabIndex        =   15
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost :"
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
      TabIndex        =   13
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pricing Information"
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
      Left            =   360
      TabIndex        =   11
      Top             =   3480
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
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
      TabIndex        =   10
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
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
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Short Description :"
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
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Long Description :"
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
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Product Item Information"
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
      TabIndex        =   2
      Top             =   1200
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "UPC/Barcode :"
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
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "frmItem.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8205
   End
End
Attribute VB_Name = "frmMMSitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbox_Click(Index As Integer)
Select Case Index
    Case 0
        Call dbconek
        With ar
            .Open "Select *From supplier Where supplierdesc='" & cbox(0) & "'", strConek, adOpenStatic, adLockOptimistic
                lblsup = !supplierid
            .Close
        End With
    
    Case 1
        Call dbconek
        With ar
            .Open "Select *From category Where categorydesc='" & cbox(1) & "'", strConek, adOpenStatic, adLockOptimistic
                lblcat = !categoryid
            .Close
        End With
End Select
End Sub

Private Sub cbox_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboxdiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboxinterest_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdEdit_Click()
    Call txt_unlock
    txt(0).SetFocus
    
    cmdnew.Enabled = False
    cmdsave.Caption = "&Update"
    cmdsave.Enabled = True
    Command1.Enabled = True
    cmdedit.Enabled = False
    cmdlookup.Enabled = True
    
    lblins = "Enter the UPC or Barcode number then press enter. Click Look Up button if you want to search by description"
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdlookup_Click()
    txt(0).SetFocus
    itemflag = 1
    LoadForm ItemSearch
End Sub

Private Sub cmdnew_Click()
    Call txt_unlock
    txt(0).SetFocus
    
    cmdnew.Enabled = True
    cmdsave.Enabled = True
    Command1.Enabled = True
    cmdedit.Enabled = False
    
    lblins = "Fill up all the fields then press save. Click cancel button to cancel"
End Sub

Private Sub cmdsave_Click()
Select Case cmdsave.Caption
    Case "&Save"
        If cbox(0) = "" Then
            MsgBox "You must select Supplier for this item", vbCritical, "Error Saving"
            cbox(0).SetFocus
            Exit Sub
        End If
        If cbox(1) = "" Then
            MsgBox "You must select category for this item", vbCritical, "Error Saving"
            cbox(0).SetFocus
            Exit Sub
        End If
        For i = 0 To 2
            If txt(i) = "" Then
                MsgBox "All fields are required", vbCritical, "Error Saving"
                txt(i).SetFocus
                Exit Sub
            End If
        Next
        If Val(txt(3)) = 0 Or Val(txt(4)) = 0 Then
            MsgBox "Pricing for this item must have a value.", vbCritical, "Error Saving"
            txt(3).SetFocus
            Exit Sub
        End If
        Call dbconek
        With ar
            .Open "Select *From itemmaster", strConek, adOpenStatic, adLockOptimistic
                .AddNew
                    !upc = txt(0)
                    !longdesc = txt(1)
                    !shortdesc = txt(2)
                    !supplierid = lblsup
                    !categoryid = lblcat
                    !unitcost = txt(3)
                    !unitprice = txt(4)
                    !discount = cboxdiscount.Text
                    !interestincredit = cboxinterest.Text
                    !stock = "0"
                    !Status = "NEW ITEM"
                .Update
            .Close
        End With
    Case "&Update"
        If cbox(0) = "" Then
            MsgBox "You must select Supplier for this item", vbCritical, "Error Saving"
            cbox(0).SetFocus
            Exit Sub
        End If
        If cbox(1) = "" Then
            MsgBox "You must select category for this item", vbCritical, "Error Saving"
            cbox(0).SetFocus
            Exit Sub
        End If
        For i = 0 To 2
            If txt(i) = "" Then
                MsgBox "All fields are required", vbCritical, "Error Saving"
                txt(i).SetFocus
                Exit Sub
            End If
        Next
        If Val(txt(3)) = 0 Or Val(txt(4)) = 0 Then
            MsgBox "Pricing for this item must have a value.", vbCritical, "Error Saving"
            txt(3).SetFocus
            Exit Sub
        End If
        Call dbconek
        With ar
            .Open "Select *From itemmaster Where upc='" & txt(0) & "'", strConek, adOpenStatic, adLockOptimistic
                If .RecordCount = 0 Then
                    MsgBox "Upc cannot be change, Add this item as new item", vbCritical, "Updating Error"
                End If
                    !longdesc = txt(1)
                    !shortdesc = txt(2)
                    !supplierid = lblsup
                    !categoryid = lblcat
                    !unitcost = txt(3)
                    !unitprice = txt(4)
                    !discount = cboxdiscount.Text
                    !interestincredit = cboxinterest.Text
                .Update
            .Close
        End With
End Select
End Sub

Private Sub Command1_Click()
    Unload Me
    LoadForm frmMMSitem
End Sub

Private Sub Form_Load()
    cboxdiscount.AddItem "0.01"
    cboxdiscount.AddItem "0.02"
    cboxdiscount.AddItem "0.03"
    cboxdiscount.AddItem "0.04"
    cboxdiscount.AddItem "0.05"
    cboxdiscount.AddItem "0.06"
    cboxdiscount.AddItem "0.07"
    cboxdiscount.AddItem "0.08"
    cboxdiscount.AddItem "0.09"
    cboxdiscount.AddItem "0.10"
    
    cboxinterest.AddItem "0.01"
    cboxinterest.AddItem "0.02"
    cboxinterest.AddItem "0.03"
    cboxinterest.AddItem "0.04"
    cboxinterest.AddItem "0.05"
    cboxinterest.AddItem "0.06"
    cboxinterest.AddItem "0.07"
    cboxinterest.AddItem "0.08"
    cboxinterest.AddItem "0.09"
    cboxinterest.AddItem "0.10"
    
    Me.Icon = frmNewEmployee.SmallImages.ListImages(1).Picture
    
    Call dbconek
    With ar
        .Open "Select *From supplier", strConek, adOpenStatic, adLockOptimistic
            .MoveFirst
            Do While Not .EOF
                cbox(0).AddItem !supplierdesc
                .MoveNext
            Loop
        .Close
    End With
    Call dbconek
    With ar
        .Open "Select *From category", strConek, adOpenStatic, adLockOptimistic
            .MoveFirst
            Do While Not .EOF
                cbox(1).AddItem !categorydesc
                .MoveNext
            Loop
        .Close
    End With
    
End Sub


Private Sub txt_GotFocus(Index As Integer)
    SendKeys "{end}+{home}"
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

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Select Case Index
    Case 0
        If KeyAscii = 13 Then
            Call dbconek
            With ar
                .Open "Select *From itemmaster Where upc='" & txt(0) & "'", strConek, adOpenStatic, adLockOptimistic
                    If .RecordCount >= 1 Then
                    txt(1) = !longdesc
                    txt(2) = !shortdesc
                    lblsup.Caption = !supplierid
                    lblcat.Caption = !categoryid
                    txt(3) = Format(CCur(!unitcost), "###,##0.00")
                    txt(4) = Format(CCur(!unitprice), "###,##0.00")
                    cboxdiscount = !discount
                    cboxinterest = !interestincredit
                    Else
                    MsgBox "Not item with that upc", vbInformation, "Search Error"
                    txt(0).SetFocus
                    Exit Sub
                    End If
                .Close
            End With
            
            Call dbconek
            ar.Open "Select *From supplier Where supplierid='" & lblsup & "'", strConek, adOpenStatic, adLockOptimistic
                cbox(0) = ar!supplierdesc
            ar.Close
            
            Call dbconek
            ar.Open "Select *From category Where categoryid='" & lblcat & "'", strConek, adOpenStatic, adLockOptimistic
                cbox(1) = ar!categorydesc
            ar.Close
        End If
    Case 1
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
End Select
End Sub
