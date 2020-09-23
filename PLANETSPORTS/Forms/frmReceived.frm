VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMMSReceived 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase  Order Receiving"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   2145
   ClientWidth     =   11205
   Icon            =   "frmReceived.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReceived.frx":000C
   ScaleHeight     =   8670
   ScaleMode       =   0  'User
   ScaleWidth      =   11200
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   7680
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   4440
      ScaleHeight     =   1305
      ScaleWidth      =   2985
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC to Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
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
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Quantity:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   9360
      TabIndex        =   2
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   " &Post"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   8040
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstpo 
      Height          =   2415
      Left            =   480
      TabIndex        =   18
      Top             =   2160
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UPC/Barcode"
         Object.Width           =   2541
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   8823
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Unit Cost"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2541
      EndProperty
   End
   Begin MSComctlLib.ListView lstrec 
      Height          =   2535
      Left            =   480
      TabIndex        =   19
      Top             =   5040
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UPC/Barcode"
         Object.Width           =   2541
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   8823
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Unit Cost"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2541
      EndProperty
   End
   Begin VB.Label lblrec 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   7440
      TabIndex        =   20
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Received Items"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   4680
      Width           =   10215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Purchase Order Items"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1800
      Width           =   10215
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   480
      Picture         =   "frmReceived.frx":2B37
      Top             =   8040
      Width           =   480
   End
   Begin VB.Label lblsup 
      Caption         =   "Label1"
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblcost 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblq 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblupc 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbldesc 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To received the items from the supplier you must have the Purchase Order Form. Type the P.O. Number the press <ENTER> key."
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   5
      Top             =   8160
      Width           =   5535
   End
   Begin VB.Label LBLPURID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   8400
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Id:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "frmReceived.frx":2F79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11205
   End
End
Attribute VB_Name = "frmMMSReceived"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstupc As String
Dim lstdesc As String
Dim lstquantity As String
Dim lstunitcost As String
Dim lsttotal As String
Dim recupc, recdesc, reccost, recq, rectotal As String
Dim INTTOTAL As Double
Private Sub cmdexit_Click()
    Unload Me
End Sub
Private Sub cmdsave_Click()
On Error Resume Next
Dim recitem As Integer
    
        INTOTAL = 0
        
        'Save to table "received"'
        'Begin:
        Set ac = New ADODB.Connection
        Set ar = New ADODB.Recordset
            
        Call dbconek
        ac.Open strConek
            
        With ar
            criteria = "Select *From received"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
                .AddNew
                    !receivedid = lblrec
                    !supplierid = lblsup
                    !datereceived = Format(Date, "mm/dd/yyyy")
                    !receivedby = currentemmsuser
                .Update
            .Close
        End With
        'End
        '***********************************************************
        
        pbar.Value = 0
        pbar.Min = 0
        pbar.Max = 100
               
        '***********************************************************
        'Save the contents of the listbox "lstrec"
        
        For recitem = 1 To lstrec.ListItems.Count  'Counter for listbox "lstrec"
            Set ac = New ADODB.Connection
            Set ar = New ADODB.Recordset
    
            Call dbconek
            ac.Open strConek
            
            With ar
                criteria = "Select *From receiveddetails"
                .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    .AddNew
                        !receivedid = lblrec.Caption
                        !upc = lstrec.ListItems(recitem).Text
                        !Description = lstrec.ListItems(recitem).SubItems(1)
                        !datereceived = Format(Date, "mm/dd/yyyy")
                        !quantity = lstrec.ListItems(recitem).SubItems(3)
                        !unitcost = lstrec.ListItems(recitem).SubItems(2)
                        !total = lstrec.ListItems(recitem).SubItems(4)
                        !supplierid = lblsup.Caption
                        !Status = "SERVED"
                    .Update
                .Close
            End With
            
            '*****************************************************
            'Update the Item Master File
            Set ac = New ADODB.Connection
            Set ar = New ADODB.Recordset
            
            Call dbconek
            ac.Open strConek
            
            With ar
                criteria = "Select *From itemmaster Where upc='" & lstrec.ListItems(recitem).Text & "'"
                .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    !stock = Val(!stock) + Val(lstrec.ListItems(recitem).SubItems(3))
                    .Update
                .Close
            End With
            'End
            
            Set ac = New ADODB.Connection
            Set ar = New ADODB.Recordset
            
            Call dbconek
            ac.Open strConek
            
            With ar
                criteria = "Select *From podetails Where poid='" & txt(0) & "' And upc='" & lstrec.ListItems(recitem).Text & "'"
                .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    !servedquantity = lstrec.ListItems(recitem).SubItems(3)
                    !Status = "SERVED"
                    .Update
                .Close
            End With
            
            If pbar.Value + (100 / lstrec.ListItems.Count) >= 100 Then
               pbar.Value = 100
            Else
               pbar.Value = pbar.Value + (100 / lstrec.ListItems.Count)
            End If
            
            lstpo.ListItems.Clear
            
            INTOTAL = INTOTAL + lstrec.ListItems(ii).SubItems(4)
        Next
            '***********************************************************
        
        
        '***********************************************************
        'Update the table "purchaseorder"'
        'Begin:
        
        Set ac = New ADODB.Connection
        Set ar = New ADODB.Recordset
            
        Call dbconek
        ac.Open strConek
            
        With ar
            criteria = "Select *From purchaseorder Where poid='" & txt(0) & "'"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
                !dateserved = Format(Date, "mm/dd/yyyy")
                .Update
            .Close
        End With
        
        A = MsgBox("Purchase posting was successful. Do you want to print the P.O.?" & vbCrLf & "If you click No it will be saved to a file then print it later..", vbYesNo + vbQuestion, "Purchase had been posted")
        If A = vbYes Then
            Printer.Print " "
            Printer.Print "RECEIVED NUMBER: " & lblrec
            Printer.Print "SUPPLIER ID: " & lblsup
            Printer.Print "DATE RECEIVED: " & Format(Date, "MM/DD/YYYY")
            Printer.Print "UPC/BARCODE" & Space(15 - Len("UPC/BARCODE")) & "LONG DESCRIPTION" & Space(40 - Len("LONG DESCRIPTION")) & "UNIT COST" & Space(15 - Len("UNIT COST")) & "QUANTITY" & Space(15 - Len("QUANTITY")) & "TOTAL COST"
            Printer.Print "----------------------------------------------------------------------------------------------------"
            For ii = 1 To lstrec.ListItems.Count
                Printer.Print lstrec.ListItems(ii).Text & Space(15 - Len(lstrec.ListItems(ii).Text)) & lstrec.ListItems(ii).SubItems(1) & Space(40 - Len(lstrec.ListItems(ii).SubItems(1))) & Format(CCur(lstrec.ListItems(ii).SubItems(2)), "###,###,##0.00") & Space(15 - Len(Format(CCur(lstrec.ListItems(ii).SubItems(2)), "###,###,##0.00"))) & Format(lstrec.ListItems(ii).SubItems(3), "###,###,##0.00") & Space(15 - Len(Format(lstrec.ListItems(ii).SubItems(3), "###,###,##0.00"))) & Format(CCur(lstrec.ListItems(ii).SubItems(4)), "###,###,##0.00")
            Next
            Printer.Print "----------------------------------------------------------------------------------------------------"
            Printer.Print ""
            Printer.Print "RECEIVED BY: " & currentemmsuser
            
        Else
            Open App.Path & "\Received.txt" For Append As #1
                Print #1, " "
                Print #1, "RECEIVED NUMBER: " & lblrec
                Print #1, "SUPPLIER ID: " & lblsup
                Print #1, "DATE RECEIVED: " & Format(Date, "MM/DD/YYYY")
                Print #1, "UPC/BARCODE" & Space(15 - Len("UPC/BARCODE")) & "LONG DESCRIPTION" & Space(40 - Len("LONG DESCRIPTION")) & "UNIT COST" & Space(15 - Len("UNIT COST")) & "QUANTITY" & Space(15 - Len("QUANTITY")) & "TOTAL COST"
                Print #1, "----------------------------------------------------------------------------------------------------"
                For ii = 1 To lstrec.ListItems.Count
                    Print #1, lstrec.ListItems(ii).Text & Space(15 - Len(lstrec.ListItems(ii).Text)) & lstrec.ListItems(ii).SubItems(1) & Space(40 - Len(lstrec.ListItems(ii).SubItems(1))) & Format(CCur(lstrec.ListItems(ii).SubItems(2)), "###,###,##0.00") & Space(15 - Len(Format(CCur(lstrec.ListItems(ii).SubItems(2)), "###,###,##0.00"))) & Format(lstrec.ListItems(ii).SubItems(3), "###,###,##0.00") & Space(15 - Len(Format(lstrec.ListItems(ii).SubItems(3), "###,###,##0.00"))) & Format(CCur(lstrec.ListItems(ii).SubItems(4)), "###,###,##0.00")
                Next
                Print #1, "----------------------------------------------------------------------------------------------------"
                Print #1, " "
                Print #1, "RECEIVED BY: " & currentemmsuser
            Close #1
    End If
        
        Unload Me
        Load frmMMSReceived
        frmMMSReceived.Show vbModal
                
        'End
        '***********************************************************
        'end
End Sub

Private Sub Form_Load()
    Me.Icon = frmNewEmployee.SmallImages.ListImages(15).Picture
End Sub

Private Sub lstpO_Click()
    Dim retim As Integer
        
End Sub

Private Sub lstpo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
    Case vbKeyEscape
        Unload Me
        Load frmMMSReceived
        frmMMSReceived.Show vbModal
End Select
End Sub

Private Sub lstpO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Picture1.Visible = True
    
        txt(0).Enabled = False
        lstpo.Enabled = False
        lstrec.Enabled = False
        cmdsave.Enabled = False
        cmdexit.Enabled = False
        
        txt(1).SetFocus
        SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Dim istat As String
Select Case Index
Case 0
    If KeyAscii = 13 Then
        
        Set ac = New ADODB.Connection
        Set ar = New ADODB.Recordset
    
        Call dbconek
        ac.Open strConek
        
        Dim lastid As String
        With ar
            criteria = "Select *From received"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
            .MoveLast
            
                lastid = Mid(!receivedid, 5, Len(!receivedid))
                lastid = Val(lastid) + 1
                
                If Len(lastid) = 1 Then
                     lblrec.Caption = "REC-" & "00000" & lastid
                ElseIf Len(lastid) = 2 Then
                    lblrec.Caption = "REC-" & "0000" & lastid
                ElseIf Len(lastid) = 3 Then
                    lblrec.Caption = "REC-" & "000" & lastid
                ElseIf Len(lastid) = 4 Then
                    lblrec.Caption = "REC-" & "00" & lastid
                ElseIf Len(lastid) = 5 Then
                    lblrec.Caption = "REC-" & "0" & lastid
                Else
                    lblrec.Caption = "REC-" & lastid
                End If
            .Close
        End With
        
        istat = "UNSERVED"
        
        Set ac = New ADODB.Connection
        Set ar = New ADODB.Recordset
    
        Call dbconek
        ac.Open strConek
        
        
        With ar
            criteria = "Select *From podetails Where poid='" & txt(0).Text & "' And status='" & istat & "'"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
                If .RecordCount >= 1 Then
                    .MoveFirst
                    Do While Not .EOF
                        lstpo.ListItems.Add , , !upc
                        lstpo.ListItems(lstpo.ListItems.Count).SubItems(1) = !Description
                        lstpo.ListItems(lstpo.ListItems.Count).SubItems(2) = !unitcost
                        lstpo.ListItems(lstpo.ListItems.Count).SubItems(3) = !orderedquantity
                        lstpo.ListItems(lstpo.ListItems.Count).SubItems(4) = !totalcost
                                                
                        .MoveNext
                    Loop
                    Label2(1) = "Use the arrow keys to select the item to receive then press <ENTER> key. Press <ESC> to Cancel this transaction."
                Else
                    MsgBox "Purchase order does not exist or was already served", vbInformation, "Search Error"
                    txt(0).SetFocus
                    SendKeys "{home}+{end}"
                    Exit Sub
                End If
                lstpo.SetFocus
                txt(0).Enabled = False
            .Close
        End With
        
        Set ac = New ADODB.Connection
        Set ar = New ADODB.Recordset
    
        Call dbconek
        ac.Open strConek
        
        With ar
            criteria = "Select *From purchaseorder Where poid='" & txt(0) & "'"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
                lblsup = !supplierid
            .Close
        End With
        
        
    End If

    KeyAscii = Asc(UCase(Chr((KeyAscii))))
    
    Case 1
        If KeyAscii = 13 Then
            If Val(txt(1)) > Val(lstpo.ListItems(lstpo.SelectedItem.Index).SubItems(3)) Then
                MsgBox "Received Quantity is Greater the the oredered quantity. Please check", vbInformation, "Quantity Error"
                txt(1).SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            Else
                
                lstrec.ListItems.Add , , lstpo.ListItems(lstpo.SelectedItem.Index).Text
                lstrec.ListItems(lstrec.ListItems.Count).SubItems(1) = lstpo.ListItems(lstpo.SelectedItem.Index).SubItems(1)
                lstrec.ListItems(lstrec.ListItems.Count).SubItems(2) = lstpo.ListItems(lstpo.SelectedItem.Index).SubItems(2)
                lstrec.ListItems(lstrec.ListItems.Count).SubItems(3) = txt(1)
                lstrec.ListItems(lstrec.ListItems.Count).SubItems(4) = lstpo.ListItems(lstpo.SelectedItem.Index).SubItems(4) * Val(txt(1))
                
                Picture1.Visible = False
                lstpo.Enabled = True
                lstrec.Enabled = True
                cmdsave.Enabled = True
                cmdexit.Enabled = True
                
                lstpo.SetFocus
            End If
            
            lstpo.ListItems.Remove (lstpo.SelectedItem.Index)
        End If
        
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        
        End If
                
End Select
End Sub
