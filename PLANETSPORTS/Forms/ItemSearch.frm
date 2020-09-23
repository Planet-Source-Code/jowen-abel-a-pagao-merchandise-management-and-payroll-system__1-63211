VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ItemSearch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product/Item Search"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2990
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Supplier Name"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Supplier ID"
         Object.Width           =   2540
      EndProperty
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
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Description:"
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
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "ItemSearch.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10620
   End
End
Attribute VB_Name = "ItemSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If itemflag = 1 Then
        frmMMSitem.txt(0) = lv.ListItems(lv.SelectedItem.Index).SubItems(1)
        Unload Me
    
    ElseIf itemflag = 2 Then
        frmMMSPurchaseOrder.txt(0) = lv.ListItems(lv.SelectedItem.Index).SubItems(1)
        Unload Me
    
    ElseIf itemflag = 3 Then
        frmReturn.txt(0) = lv.ListItems(lv.SelectedItem.Index).SubItems(1)
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Image1.Width = Me.Width
    Me.Icon = SupplierSearch.img.ListImages(3).Picture
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dbconek
        
        With ar
            criteria = "Select * From itemmaster Where longdesc like'" & txt(0) & "%'"
            .Open criteria, strConek, 3, 3
                If .RecordCount >= 1 Then
                    lv.ListItems.Clear
                Do While Not .EOF
                    lv.ListItems.Add , , !longdesc
                    lv.ListItems(lv.ListItems.Count).SubItems(1) = !upc
                    .MoveNext
                Loop
                lv.SetFocus
                Else
                    MsgBox "Item does not exist.", vbInformation, "Search Error"
                    txt(0).SetFocus
                    Exit Sub
                End If
            .Close
        End With
    End If
End Sub
