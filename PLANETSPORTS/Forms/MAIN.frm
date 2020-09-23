VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MAIN 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Employee and Merchandise Management System V 1.0     jowen_abel@yahoo.com"
   ClientHeight    =   8760
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12090
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   12030
      TabIndex        =   1
      Top             =   780
      Width           =   12090
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8460
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
            Picture         =   "MAIN.frx":0000
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "User Name:"
            TextSave        =   "User Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7294
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "MAIN.frx":039C
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2/28/2002"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "1:04 AM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   720
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":0736
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1410
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2262
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":30B4
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":398E
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4268
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4B42
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":550C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":5DE6
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":6100
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":69DA
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":72B4
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7B8E
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7EA8
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":8782
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":905C
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":9936
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":A210
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":AAEA
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":B3C4
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":BC9E
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":C578
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":CE52
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":D72C
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":E006
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":E8E0
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":F1BA
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":FA94
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1036E
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":10C48
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":114FE
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":11DD8
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1222A
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1267C
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":14E2E
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   120
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":160B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":17A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":193D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1AD66
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1C6F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1E08A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1FA1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":213AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":22D40
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":246D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":253B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":25C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2696C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":27648
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":28324
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":29000
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":29CDC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1376
      ButtonWidth     =   2355
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Shortcuts"
            Key             =   "Shortcuts"
            Object.ToolTipText     =   "Ctrl+F1"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Returned History"
            Key             =   "New"
            Object.ToolTipText     =   "Ctrl+F2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Purchase History"
            Key             =   "Edit"
            Object.ToolTipText     =   "Ctrl+F3"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Received History"
            Key             =   "Search"
            Object.ToolTipText     =   "Ctrl+F4"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Out of Stock"
            Key             =   "Delete"
            Object.ToolTipText     =   "Ctrl+F5"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Payroll History"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Ctrl+F6"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Loans"
            Key             =   "Print"
            Object.ToolTipText     =   "Ctrl+F7"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Close"
            Object.ToolTipText     =   "Ctrl+F8"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Menu F 
      Caption         =   "&File"
      Begin VB.Menu NI 
         Caption         =   "New &Item/Product"
         Shortcut        =   ^I
      End
      Begin VB.Menu NC 
         Caption         =   "New &Category"
         Shortcut        =   ^C
      End
      Begin VB.Menu NS 
         Caption         =   "New &Supplier"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu trans 
      Caption         =   "&Transactions"
      Begin VB.Menu p 
         Caption         =   "Pay&roll"
         Shortcut        =   ^S
      End
      Begin VB.Menu RI 
         Caption         =   "Returned &Items"
         Shortcut        =   ^R
      End
      Begin VB.Menu RO 
         Caption         =   "&Receive Order"
         Shortcut        =   ^O
      End
      Begin VB.Menu PO 
         Caption         =   "&Purchase Order"
         Shortcut        =   ^P
      End
      Begin VB.Menu el 
         Caption         =   "Employee &Loan"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu r 
      Caption         =   "&Reports"
      Begin VB.Menu SIR 
         Caption         =   "&Sales Income Report"
         Shortcut        =   {F5}
      End
      Begin VB.Menu rr 
         Caption         =   "&Received Report"
         Shortcut        =   {F6}
      End
      Begin VB.Menu PR 
         Caption         =   "&Purchase Report"
         Shortcut        =   {F7}
      End
      Begin VB.Menu RIR 
         Caption         =   "Returned &Items Report"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu U 
      Caption         =   "&Utilities"
      Begin VB.Menu NE 
         Caption         =   "&New Employee"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu NU 
         Caption         =   "New &User"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu BUD 
         Caption         =   "&Back Up Database"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu RD 
         Caption         =   "&Restore Database"
         Shortcut        =   ^{F6}
      End
   End
   Begin VB.Menu t 
      Caption         =   "&Tables"
      Begin VB.Menu wht 
         Caption         =   "&With Holding Tax"
         Shortcut        =   ^W
      End
      Begin VB.Menu ph 
         Caption         =   "&Phil. Health"
         Shortcut        =   ^H
      End
      Begin VB.Menu ss 
         Caption         =   "&Social Security"
         Shortcut        =   ^Y
      End
      Begin VB.Menu rap 
         Caption         =   "&Rate and Position"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu A 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A_Click()
    LoadForm frmAbout
End Sub

Private Sub BUD_Click()
    LoadForm FRM_BACK_UP
End Sub

Private Sub el_Click()
    LoadForm Form1
End Sub

Private Sub MDIForm_Load()
    
    
    
    Me.Icon = frmShortcuts.ImageList1.ListImages(5).Picture
       
    
    Me.StatusBar1.Panels(3).Text = currentemmsuser
    Me.StatusBar1.Panels(4).Text = CurrentPosition
End Sub

Private Sub NC_Click()
    LoadForm frmMMSCategory
End Sub

Private Sub NE_Click()
    LoadForm frmNewEmployee
End Sub

Private Sub NI_Click()
    LoadForm frmMMSitem
End Sub

Private Sub NS_Click()
    LoadForm frmMMSSupplier
End Sub

Private Sub NU_Click()
    LoadForm frmUser
End Sub

Private Sub p_Click()
    LoadForm frmPayroll
End Sub

Private Sub ph_Click()
    LoadForm frmPH
End Sub

Private Sub PO_Click()
    LoadForm frmMMSPurchaseOrder
End Sub

Private Sub PR_Click()
    LoadForm ReportPurchased
End Sub

Private Sub rap_Click()
    LoadForm frmRate
End Sub

Private Sub RD_Click()
    LoadForm FRM_RESTORE
End Sub

Private Sub RI_Click()
    LoadForm frmReturn
End Sub

Private Sub RO_Click()
    LoadForm frmMMSReceived
End Sub

Private Sub RR_Click()
    LoadForm ReportDelivery
End Sub

Private Sub SIR_Click()
    LoadForm ReportSales
End Sub

Private Sub ss_Click()
    LoadForm frmSSS
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
Select Case Button.Index
    Case 1: frmShortcuts.Show
    Case 3
        historyflag = 4
        Unload frmHistoryReports
        Load frmHistoryReports
        frmHistoryReports.Show
    Case 4
        historyflag = 3
        Unload frmHistoryReports
        Load frmHistoryReports
        frmHistoryReports.Show
    Case 5
        historyflag = 2
        Unload frmHistoryReports
        Load frmHistoryReports
        frmHistoryReports.Show
    Case 6: frmos.Show
    Case 7
        historyflag = 1
        Unload frmHistoryReports
        Load frmHistoryReports
        frmHistoryReports.Show
    Case 8
        Open App.Path & "\Loan.txt" For Output As #1
                        Print #1, ""
                    Close #1
                    
        historyflag = 5
                    
                
        
        Call dbconek
        With ar
            .Open "Select *From loantable", strConek, adOpenStatic, adLockOptimistic
                .MoveFirst
                Open App.Path & "\Loan.txt" For Append As #1
                    Print #1, "Loan ID" & Space(15 - Len("Loan ID")) & "Emp ID" & Space(15 - Len("Emp ID")) & "Employee Name" & Space(40 - Len("Employee Name")) & "Loan Amount" & Space(15 - Len("Loan Amount")) & "Term" & Space(15 - Len("Term")) & "Deduction" & Space(15 - Len("Deduction")) & "Paid"
                    Print #1, ""
                    Print #1, "-----------------------------------------------------------------------------------------------------------------------------"
                Close #1
                Do While Not .EOF
                    Open App.Path & "\Loan.txt" For Append As #1
                    Print #1, !loannumber & Space(15 - Len(!loannumber)) & !employeeid & Space(15 - Len(!employeeid)) & !employeename & Space(40 - Len(!employeename)) & Format(CCur(!loanamount), "###,###,##0.00") & Space(15 - Len(Format(CCur(!loanamount), "###,###,##0.00"))) & !loanterm & Space(15 - Len(!loanterm)) & Format(CCur(!deductionpermonth), "###,###,##0.00") & Space(15 - Len(Format(CCur(!deductionpermonth), "###,###,##0.00"))) & !paid
                    Close #1
                    .MoveNext
                Loop
                
            .Close
        End With
        
        Unload frmHistoryReports
        Load frmHistoryReports
        frmHistoryReports.Show
    Case 9: Unload Me
        
End Select
End Sub




Private Sub wht_Click()
    LoadForm frmTAX
End Sub
