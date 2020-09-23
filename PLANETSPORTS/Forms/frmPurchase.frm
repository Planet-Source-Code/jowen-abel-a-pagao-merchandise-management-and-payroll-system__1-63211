VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMMSPurchaseOrder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   2145
   ClientWidth     =   11205
   Icon            =   "frmPurchase.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleMode       =   0  'User
   ScaleWidth      =   11200
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   6480
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdremove 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   9120
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdlookup 
      Caption         =   "&Look Up"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   6480
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1560
      Width           =   4215
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
      Height          =   285
      Index           =   2
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2760
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
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2760
      Width           =   5295
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.ComboBox cbox 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
   End
   Begin MSComctlLib.ListView lst 
      Height          =   2895
      Left            =   480
      TabIndex        =   9
      Top             =   3480
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
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
   Begin VB.Label lblsup 
      Caption         =   "Label3"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost:"
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
      Left            =   7920
      TabIndex        =   23
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11155.02
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label lblins 
      BackStyle       =   0  'Transparent
      Caption         =   "For new purchase click &New"
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   7200
      Width           =   10215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Instruction"
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
      Index           =   4
      Left            =   480
      TabIndex        =   21
      Top             =   6840
      Width           =   10215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Agreement Information"
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
      Index           =   3
      Left            =   5760
      TabIndex        =   20
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Memo:"
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
      Left            =   5760
      TabIndex        =   19
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Quantity:"
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
      Left            =   9120
      TabIndex        =   18
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Long Description:"
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
      Left            =   2520
      TabIndex        =   16
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter UPC/Barcode:"
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
      Left            =   480
      TabIndex        =   15
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Supplier :"
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
      Left            =   480
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Product/Item Information"
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
      Left            =   480
      TabIndex        =   12
      Top             =   2160
      Width           =   10215
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
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "List of Products/Items Ordered"
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
      Left            =   480
      TabIndex        =   10
      Top             =   3120
      Width           =   10215
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "frmPurchase.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11205
   End
End
Attribute VB_Name = "frmMMSPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id As String
Private Sub cbox_Click()
On Error Resume Next
    Call dbconek
    
    ar.Open "Select *From supplier Where supplierdesc='" & cbox & "'", strConek, adOpenStatic, adLockOptimistic
        lblsup = ar!supplierid
    ar.Close
    
    txt(3).SetFocus
End Sub

Private Sub cbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDown Or KeyAscii = vbKeyUp Then
        Exit Sub
    End If
    
    KeyAscii = 0
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdlookup_Click()
    txt(0).SetFocus
    itemflag = 2
    LoadForm ItemSearch
End Sub

Private Sub cmdnew_Click()
On Error Resume Next

Dim aa
        Call dbconek
        
        Dim lastid As String
        With ar
            criteria = "Select *From purchaseorder"
            .Open criteria, strConek, adOpenStatic, adLockOptimistic
            .MoveLast
            
                lastid = Mid(!poid, 5, Len(!poid))
                lastid = Val(lastid) + 1
                
                If Len(lastid) = 1 Then
                    Me.Caption = "Purchase Order - " & "PUR-" & "00000" & lastid
                ElseIf Len(lastid) = 2 Then
                    Me.Caption = "Purchase Order - " & "PUR-" & "0000" & lastid
                ElseIf Len(lastid) = 3 Then
                    Me.Caption = "Purchase Order - " & "PUR-" & "000" & lastid
                ElseIf Len(lastid) = 4 Then
                    Me.Caption = "Purchase Order - " & "PUR-" & "00" & lastid
                ElseIf Len(lastid) = 5 Then
                    Me.Caption = "Purchase Order - " & "PUR-" & "0" & lastid
                Else
                    Me.Caption = "Purchase Order - " & "PUR-" & lastid
                End If
            .Close
        End With
        
        id = Mid(Me.Caption, 18, 15)
        
        cbox.Enabled = True
        txt(0).Enabled = True
        txt(3).Enabled = True
        txt(4).Enabled = True
        Command1.Enabled = True
        cmdsave.Enabled = True
        cmdlookup.Enabled = True
        cmdnew.Enabled = False
        lblins = "Please select supplier and type your agreement with this P.O. between the store and the supplier."
        
        For aa = 0 To 4
            txt(aa) = ""
        Next
        cbox = ""
        cbox.SetFocus
End Sub

Private Sub cmdremove_Click()
    If lst.ListItems.Count = 0 Then
        Exit Sub
    End If
    lst.ListItems.Remove (lst.SelectedItem.Index)
End Sub

Private Sub cmdsave_Click()
Dim i, ii As Integer
Dim A
On Error Resume Next
    If lst.ListItems.Count = 0 Then
        MsgBox "There is nothing to save", vbInformation, "Saving Error"
        Exit Sub
    End If
    
    'Save to parent table 1 to many relatonship
    Call dbconek
    With ar
        .Open "Select *From purchaseorder", strConek, adOpenStatic, adLockOptimistic
            .AddNew
                !poid = id
                !podate = Format(Date, "mm/dd/yyyy")
                !supplierid = lblsup
                !supplierdesc = cbox.Text
                !purchaseby = currentemmsuser
                !Memo = txt(3)
                !printed = "NO"
            .Update
        .Close
    End With
    
    pbar.Min = 0
    pbar.Max = 100
    pbar.Value = 0
    
    Call dbconek
    For i = 1 To lst.ListItems.Count
        With ar
            .Open "Select *From podetails", strConek, adOpenStatic, adLockOptimistic
                .AddNew
                    !poid = id
                    !upc = lst.ListItems(i).Text
                    !Description = lst.ListItems(i).SubItems(1)
                    !unitcost = lst.ListItems(i).SubItems(2)
                    !orderedquantity = lst.ListItems(i).SubItems(3)
                    !totalcost = lst.ListItems(i).SubItems(4)
                    !supplierid = lblsup
                    !Status = "UNSERVED"
                .Update
            .Close
        End With
        
        If pbar.Value + (100 / lst.ListItems.Count) >= 100 Then
            pbar.Value = 100
        Else
            pbar.Value = pbar.Value + (100 / lst.ListItems.Count)
        End If
    Next
     
    A = MsgBox("Purchase posting was successful. Do you want to print the P.O.?" & vbCrLf & "If you click No it will be saved to a file then print it later..", vbYesNo + vbQuestion, "Purchase had been posted")
    If A = vbYes Then
        Printer.Print ""
        Printer.Print "PURCHASE ORDER NUMBER: " & id
        Printer.Print "SUPPLIER ID: " & lblsup
        Printer.Print "SUPPLIER NAME: " & cbox.Text
        Printer.Print "DATE: " & Format(Date, "MM/DD/YYYY")
        Printer.Print "UPC/BARCODE" & Space(15 - Len("UPC/BARCODE")) & "LONG DESCRIPTION" & Space(40 - Len("LONG DESCRIPTION")) & "UNIT COST" & Space(15 - Len("UNIT COST")) & "QUANTITY" & Space(15 - Len("QUANTITY")) & "TOTAL COST"
        Printer.Print "----------------------------------------------------------------------------------------------------"
        For ii = 1 To lst.ListItems.Count
            Printer.Print lst.ListItems(ii).Text & Space(15 - Len(lst.ListItems(ii).Text)) & lst.ListItems(ii).SubItems(1) & Space(40 - Len(lst.ListItems(ii).SubItems(1))) & Format(CCur(lst.ListItems(ii).SubItems(2)), "###,###,##0.00") & Space(15 - Len(Format(CCur(lst.ListItems(ii).SubItems(2)), "###,###,##0.00"))) & Format(CCur(lst.ListItems(ii).SubItems(3)), "###,###,##0.00") & Space(15 - Len(Format(CCur(lst.ListItems(ii).SubItems(3)), "###,###,##0.00"))) & Format(CCur(lst.ListItems(ii).SubItems(4)), "###,###,##0.00")
        Next
        Printer.Print "----------------------------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.Print "PURCHASED BY: " & currentemmsuser
        Printer.Print "NOTE: "
        Printer.Print "      " & txt(3)
    Else
        Open App.Path & "\PO.txt" For Append As #1
            Print #1, " "
            Print #1, "PURCHASE ORDER NUMBER: " & id
            Print #1, "SUPPLIER ID: " & lblsup
            Print #1, "SUPPLIER NAME: " & cbox.Text
            Print #1, "DATE: " & Format(Date, "MM/DD/YYYY")
            Print #1, "UPC/BARCODE" & Space(15 - Len("UPC/BARCODE")) & "LONG DESCRIPTION" & Space(40 - Len("LONG DESCRIPTION")) & "UNIT COST" & Space(15 - Len("UNIT COST")) & "QUANTITY" & Space(15 - Len("QUANTITY")) & "TOTAL COST"
            Print #1, "----------------------------------------------------------------------------------------------------"
            For ii = 1 To lst.ListItems.Count
                Print #1, lst.ListItems(ii).Text & Space(15 - Len(lst.ListItems(ii).Text)) & lst.ListItems(ii).SubItems(1) & Space(40 - Len(lst.ListItems(ii).SubItems(1))) & Format(CCur(lst.ListItems(ii).SubItems(2)), "###,###,##0.00") & Space(15 - Len(Format(CCur(lst.ListItems(ii).SubItems(2)), "###,###,##0.00"))) & Format(CCur(lst.ListItems(ii).SubItems(3)), "###,###,##0.00") & Space(15 - Len(Format(CCur(lst.ListItems(ii).SubItems(3)), "###,###,##0.00"))) & Format(CCur(lst.ListItems(ii).SubItems(4)), "###,###,##0.00")
            Next
            Print #1, "----------------------------------------------------------------------------------------------------"
            Print #1, " "
            Print #1, "PURCHASED BY: " & currentemmsuser
            Print #1, "NOTE: "
            Print #1, "      " & txt(3)
        Close #1
        
        Call Command1_Click
    End If
      
End Sub

Private Sub Command1_Click()
    Unload Me
    LoadForm frmMMSPurchaseOrder
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Icon = frmNewEmployee.SmallImages.ListImages(4).Picture
    
    Call dbconek
    
    ar.Open "Select *From supplier", strConek, adOpenStatic, adLockOptimistic
        ar.MoveFirst
        Do While Not ar.EOF
            cbox.AddItem ar!supplierdesc
            ar.MoveNext
        Loop
    ar.Close
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
                criteria = "Select *From itemmaster Where upc='" & txt(0) & "'"
                .Open criteria, strConek, adOpenStatic, adLockOptimistic
                    If .RecordCount >= 1 Then
                        If !supplierid = lblsup.Caption Then
                            txt(1) = !longdesc
                            txt(2) = Format(!unitcost, "###,###,##0.00")
                            txt(4).SetFocus
                        Else
                            MsgBox "This Item/Product is not coveredby this supplier.", vbCritical, "Error On Item"
                            txt(0).SetFocus
                            Exit Sub
                        End If
                    Else
                        MsgBox "This Item/Product is not on the database.", vbCritical, "Error Search"
                        txt(0).SetFocus
                        SendKeys "{end}+{home}"
                        Exit Sub
                    End If
                .Close
            End With
        End If
    Case 3
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 4
        If KeyAscii = 13 Then
            lst.ListItems.Add , , txt(0)
            lst.ListItems(lst.ListItems.Count).SubItems(1) = txt(1)
            lst.ListItems(lst.ListItems.Count).SubItems(2) = txt(2)
            lst.ListItems(lst.ListItems.Count).SubItems(3) = txt(4)
            lst.ListItems(lst.ListItems.Count).SubItems(4) = Format(CCur(txt(2) * txt(4)), "###,###,##0.00")
            
        txt(2) = ""
        txt(1) = ""
        txt(0) = ""
        txt(4) = ""
        txt(0).SetFocus
        cmdremove.Enabled = True
        End If
        If KeyAscii = 8 Or KeyAscii = 46 Then
            Exit Sub
        End If
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
        
        
        
End Select
End Sub
