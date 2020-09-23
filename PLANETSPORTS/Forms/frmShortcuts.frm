VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmShortcuts 
   BackColor       =   &H80000005&
   Caption         =   "Shortcuts"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShortcuts.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7350
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1725
      Top             =   4275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":2394
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":3070
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":4A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":6394
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":7D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":96B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":A392
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":B06C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":BD46
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":CA22
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":D6FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":DFDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":ECB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":F992
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1066E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":10F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":11C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1250A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":131E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":14B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1650E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvMenu 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   6376
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      MousePointer    =   99
      MouseIcon       =   "frmShortcuts.frx":16DEA
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   360
      Top             =   4200
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
            Picture         =   "frmShortcuts.frx":17AC4
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1879E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":195F0
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1A442
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1AD1C
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1B5F6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1BED0
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1C89A
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1D174
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1D48E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1DD68
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1E642
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1EF1C
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1F236
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1FB10
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":203EA
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":20CC4
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":2159E
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":21E78
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":22752
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":2302C
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":23906
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":241E0
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":24ABA
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":25394
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":25C6E
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":26548
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":26E22
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":276FC
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":27FD6
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":2888C
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":29166
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":295B8
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":29A0A
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":2C1BC
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    With lvMenu
        Set .SmallIcons = ImageList1
        Set .Icons = ImageList1
        'For Manager Shortcut
        
        If CurrentPosition = "MANAGER" Then
        .ListItems.Add , "NewEmployee", "New Employee", 1, 1
        .ListItems.Add , "NewUser", "User Maintenance", 2, 2
        .ListItems.Add , "BackUp", "Back Up Database", 8, 8
        .ListItems.Add , "Restore", "Restore Database", 9, 9
        
        'For Store Operation
        .ListItems.Add , "NewSupplier", "Supplier Control", 3, 3
        .ListItems.Add , "NewCategory", "Category Control", 5, 5
        .ListItems.Add , "NewItem", "Product Item Control", 6, 6
        .ListItems.Add , "Purchase", "Purchase Order", 15, 15
        .ListItems.Add , "Return", "Return Item", 7, 7
        .ListItems.Add , "Receive", "Warehouse Receiving", 10, 10
        
        'for reports
        
        .ListItems.Add , "Delivery", "Received Report", 12, 12
        .ListItems.Add , "ReportPurchase", "Purchase Order Report", 13, 13
        .ListItems.Add , "SalesReport", "Sales Report", 16, 16
        .ListItems.Add , "SalesIncome", "Sales Income Report", 14, 14
        
        'for payroll
        
        .ListItems.Add , "Payroll", "Create Payroll", 19, 19
        .ListItems.Add , "Rate", "Rate and Position", 22, 22
        .ListItems.Add , "Loan", "Company Loans", 20, 20
        Else
        'For Store Operation
        .ListItems.Add , "NewSupplier", "Supplier Control", 3, 3
        .ListItems.Add , "NewCategory", "Category Control", 5, 5
        .ListItems.Add , "NewItem", "Product Item Control", 6, 6
        .ListItems.Add , "Purchase", "Purchase Order", 15, 15
        .ListItems.Add , "Return", "Return Item", 7, 7
        .ListItems.Add , "Receive", "Warehouse Receiving", 10, 10
        
        'for reports
        
        .ListItems.Add , "Delivery", "Received Report", 12, 12
        .ListItems.Add , "ReportPurchase", "Purchase Order Report", 13, 13
        .ListItems.Add , "SalesReport", "Sales Report", 16, 16
        .ListItems.Add , "SalesIncome", "Sales Income Report", 14, 14
        
        'for payroll
        
        .ListItems.Add , "Payroll", "Create Payroll", 19, 19
        .ListItems.Add , "Rate", "Rate and Position", 22, 22
        .ListItems.Add , "Loan", "Company Loans", 20, 20
        End If
    End With
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    lvMenu.Width = ScaleWidth
    lvMenu.Height = ScaleHeight
End Sub




Private Sub lvMenu_DblClick()
Select Case lvMenu.SelectedItem.Key
    Case "NewEmployee": LoadForm frmNewEmployee
    Case "NewUser": LoadForm frmUser
    Case "BackUp": LoadForm FRM_BACK_UP
    Case "Restore": LoadForm FRM_RESTORE
    
    Case "NewSupplier": LoadForm frmMMSSupplier
    Case "NewCategory": LoadForm frmMMSCategory
    Case "NewItem": LoadForm frmMMSitem
    Case "Purchase": LoadForm frmMMSPurchaseOrder
    Case "Return": LoadForm frmReturn
    Case "Receive": LoadForm frmMMSReceived
    
    Case "Rate": LoadForm frmRate
    Case "Delivery": LoadForm ReportDelivery
    Case "ReportPurchase": LoadForm ReportPurchased
    Case "SalesReport": LoadForm ReportSales
    Case "Payroll": LoadForm frmPayroll
    Case "Loan": LoadForm Form1
    
End Select
End Sub
