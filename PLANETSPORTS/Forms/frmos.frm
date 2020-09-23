VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Out of Stock Items"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   2145
   ClientWidth     =   11940
   Icon            =   "frmos.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleMode       =   0  'User
   ScaleWidth      =   11934.67
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lst 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10610
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
         Text            =   "Supplier Code"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Category Code"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Stock"
         Object.Width           =   2541
      EndProperty
   End
End
Attribute VB_Name = "frmos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
On Error Resume Next

    
    Me.Icon = frmNewEmployee.SmallImages.ListImages(6).Picture
    
    Call dbconek
    
    
    With ar
        criteria = "Select *From itemmaster Where stock='" & "0" & "'"
        .Open criteria, strConek, 3, 3
            Do While Not .EOF
            lst.ListItems.Add , , !upc
            lst.ListItems(lst.ListItems.Count).SubItems(1) = !longdesc
            lst.ListItems(lst.ListItems.Count).SubItems(2) = !supplierid
            lst.ListItems(lst.ListItems.Count).SubItems(3) = !categoryid
            lst.ListItems(lst.ListItems.Count).SubItems(4) = Format(!stock, "###,###,##0.00")
            .MoveNext
            Loop
        .Close
    End With
End Sub

Private Sub Form_Resize()
    lst.Width = ScaleWidth
lst.Height = ScaleHeight
End Sub
