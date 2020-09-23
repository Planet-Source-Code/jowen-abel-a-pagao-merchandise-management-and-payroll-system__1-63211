VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ReportDelivery 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receive Report"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "ReportDelivery.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txt2 
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txt1 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cboxsupplier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   2
      Text            =   "<Select Supplier>"
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CheckBox chksupplier 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sales Report by Supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   840
      Top             =   3000
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
            Picture         =   "ReportDelivery.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":199E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":3330
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":4CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":6654
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":7FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":9978
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":B30A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":CC9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":E630
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":F30C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":FBEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":108C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":115A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":12280
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":12F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportDelivery.frx":13C38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblcat 
      Caption         =   "Label6"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblsup 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Date Range::"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   0
      Picture         =   "ReportDelivery.frx":14514
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7700
   End
End
Attribute VB_Name = "ReportDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboxcategory_Click()
On Error Resume Next
    Set ac = New ADODB.Connection
    Set ar = New ADODB.Recordset
    
    Call dbconek
    ac.Open strConek
    
    
    With ar
        criteria = "Select *From category Where categorydesc='" & cboxcategory.Text & "'"
        .Open criteria, strConek, adOpenStatic, adLockOptimistic
            lblcat.Caption = !categoryid
        .Close
    End With
End Sub

Private Sub cboxsupplier_Click()
On Error Resume Next
    Set ac = New ADODB.Connection
    Set ar = New ADODB.Recordset
    
    Call dbconek
    ac.Open strConek
    
    
    With ar
        criteria = "Select *From supplier Where supplierdesc='" & cboxsupplier.Text & "'"
        .Open criteria, strConek, adOpenStatic, adLockOptimistic
            lblsup.Caption = !supplierid
        .Close
    End With
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim rtotal As Currency
    Set ac = New ADODB.Connection
    Set ar = New ADODB.Recordset
    
    Call dbconek
    ac.Open strConek
    
    
    With ar
        criteria = "Select *From receiveddetails Where  (datereceived >= #" & SQLDate(txt1.Text) & "#) And (datereceived <= #" & SQLDate(txt2.Text) & "#)"
        .Open criteria, strConek, adOpenStatic, adLockOptimistic
            If .RecordCount > 1 Then
                If chksupplier.Value = 0 Then
                    Open App.Path & "\Report.txt" For Append As #1
                        Print #1, "Planet Sports LCCBranch"
                        Print #1, "Report Title: Delivery Report By Date Only"
                        Print #1, "Date Range: " & txt1.Text; " To " & txt2.Text
                        Print #1, "Supplier: All"
                        Print #1,
                        Print #1,
                        Print #1, "UPC/Barcode" & Space(15 - Len("UPC/Barcode")) & "Description" & Space(50 - Len("Description")) & "Unit Price" & Space(15 - Len("Unit Price")) & "Quantity" & Space(10 - Len("Quantity")) & "Total"
                        Print #1,
                        .MoveFirst
                        Do While Not .EOF
                            rtotal = rtotal + CCur(!total)
                            Print #1, !upc & Space(15 - Len(!upc)) & !Description & Space(50 - Len(!Description)) & !unitcost & Space(15 - Len(!unitcost)) & !quantity & Space(10 - Len(!quantity)) & !total
                            .MoveNext
                        Loop
                        Print #1, Space(98 - Len("Total Sales: " & Format(CCur(rtotal), "###,###,##0.00"))) & ("Total Sales: " & Format(CCur(rtotal), "###,###,##0.00"))
                    Close #1
                ElseIf chksupplier.Value = 1 Then
                        'If !supplierid = lblsup.Caption Then
                            Open App.Path & "\Report.txt" For Append As #1
                                Print #1, "Planet Sports LCCBranch"
                                Print #1, "Report Title: Delivery Report By Supplier"
                                Print #1, "Date Range: " & txt1.Text; " To " & txt2.Text
                                Print #1, "Supplier: " & cboxsupplier.Text
                                Print #1,
                                Print #1,
                                Print #1, "UPC/Barcode" & Space(15 - Len("UPC/Barcode")) & "Description" & Space(50 - Len("Description")) & "Unit Price" & Space(15 - Len("Unit Price")) & "Quantity" & Space(10 - Len("Quantity")) & "Total"
                                Print #1,
                                .MoveFirst
                                Do While Not .EOF
                                    If !supplierid = lblsup.Caption Then
                                    rtotal = rtotal + CCur(!total)
                                    Print #1, !upc & Space(15 - Len(!upc)) & !Description & Space(50 - Len(!Description)) & !unitcost & Space(15 - Len(!unitcost)) & !quantity & Space(10 - Len(!quantity)) & !total
                                    .MoveNext
                                    Else
                                        .MoveNext
                                    End If
                                Loop
                                Print #1, Space(98 - Len("Total Sales: " & Format(CCur(rtotal), "###,###,##0.00"))) & ("Total Sales: " & Format(CCur(rtotal), "###,###,##0.00"))
                            Close #1
                        'Else
                        '    MsgBox "No sales for that supplier in the given date", vbCritical, "Error"
                        '    cboxsupplier.SetFocus
                        '    Exit Sub
                        'End If
                
                End If
                Load frmReports
                frmReports.Show vbModal
            Else
                MsgBox "No Sales On Date RangeYou Ask", vbCritical, "Date Range Error"
                txt1.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
                
            End If
        .Close
        Unload Me
    End With
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call clearreport
    
    Me.Icon = Me.itb32x32.ListImages(1).Picture
    
    txt1 = Format(Date, "mm/dd/yyyy")
    txt2 = Format(Date, "mm/dd/yyyy")
    Set ac = New ADODB.Connection
    Set ar = New ADODB.Recordset
    
    Call dbconek
    ac.Open strConek
    
    With ar
        criteria = "Select *From supplier"
        .Open criteria, strConek, adOpenStatic, adLockOptimistic
            .MoveFirst
            Do While Not .EOF
                cboxsupplier.AddItem !supplierdesc
                .MoveNext
            Loop
        .Close
    End With
    
End Sub

Private Sub txt2_GotFocus()
    SendKeys "{home}+{end}"
End Sub
