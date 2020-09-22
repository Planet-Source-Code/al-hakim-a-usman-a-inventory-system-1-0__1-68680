VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Out"
   ClientHeight    =   7350
   ClientLeft      =   225
   ClientTop       =   -90
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      Height          =   975
      Left            =   -240
      TabIndex        =   21
      Top             =   -240
      Width           =   9135
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK OUT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   480
         Left            =   420
         TabIndex        =   23
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK OUT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   300
         TabIndex        =   22
         Top             =   360
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   7800
      TabIndex        =   20
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdStockout 
      Caption         =   "Stock Out"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   19
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   5400
      TabIndex        =   18
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   8775
      Begin MSComctlLib.ListView lvStockOut 
         Height          =   1935
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product ID"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product Name"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Released"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8775
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   495
         Left            =   7440
         TabIndex        =   26
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox cmbProdName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   25
         Top             =   720
         Width           =   3855
      End
      Begin VB.ComboBox cmbProdID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtSupplier 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtCategory 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtPrice 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtStocks 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtQuantity 
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6240
         TabIndex        =   1
         Top             =   2520
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DPDate_Released 
         Height          =   375
         Left            =   6600
         TabIndex        =   29
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   38753
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date Released"
         Height          =   375
         Left            =   5400
         TabIndex        =   30
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblSO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   27
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Product ID"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Supplier"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Category"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Price"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock(s)"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock Out No."
         Height          =   375
         Left            =   5400
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quantity"
         Height          =   375
         Left            =   5400
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000C&
      Height          =   2175
      Left            =   240
      TabIndex        =   16
      Top             =   4440
      Width           =   8775
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000C&
      Height          =   3135
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   8775
   End
End
Attribute VB_Name = "frmStockout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbProdID_Click()
Dim tmpID As String
tmpID = cmbProdID.Text
rs.Open "select*from tblProduct where Product_ID='" & tmpID & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    cmbProdName.Text = rs!Product_Name
    txtSupplier.Text = rs!supplier
    txtCategory.Text = rs!category
    txtPrice.Text = rs!Unit_Price
    txtStocks.Text = rs!Unit_In_Stock
End If
Set rs = Nothing
cmdAdd.Enabled = True
End Sub

Private Sub cmbProdName_Click()
Dim tmpNme As String
tmpNme = cmbProdName.Text
rs.Open "select*from tblProduct where Product_Name='" & tmpNme & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    cmbProdID.Text = rs!Product_ID
    txtSupplier.Text = rs!supplier
    txtCategory.Text = rs!category
    txtPrice.Text = rs!Unit_Price
    txtStocks.Text = rs!Unit_In_Stock
End If
Set rs = Nothing
cmdAdd.Enabled = True

End Sub

Private Sub cmdAdd_Click()
Dim siPadd, siName As String
Dim siQty, siAmount, siPrice As Double
siName = cmbProdName.Text
siQty = Val(txtQuantity.Text)
siPrice = Val(txtPrice.Text)
siAmount = siQty * siPrice
siPadd = cmbProdID.Text

Field_Check.Empty_Checks Me
If iTerminate = True Then Exit Sub

With lvStockOut
    .ListItems.Add , , siPadd
    .ListItems(.ListItems.Count).ListSubItems.Add , , siName
    .ListItems(.ListItems.Count).ListSubItems.Add , , siQty
    .ListItems(.ListItems.Count).ListSubItems.Add , , siAmount
    .ListItems(.ListItems.Count).ListSubItems.Add , , DPDate_Released.Value
End With
txtQuantity.Text = ""
cmbProdID.SetFocus
cmdStockout.Enabled = True
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
i_Clear.cLearMe Me
i_Enable.Enable_Txt Me
cmdStockout.Enabled = True
End Sub

Private Sub cmdRemove_Click()
'This will remove stockout Items list.
Dim i As Integer
For i = lvStockOut.ListItems.Count To 1 Step -1
    If lvStockOut.ListItems(i).Checked = True Then
        lvStockOut.ListItems.Remove i
    End If
Next i
If lvStockOut.ListItems.Count = 0 Then cmdStockout.Enabled = False
End Sub

Private Sub cmdStockout_Click()
ctr = 0
Dim soTmp, soTmpSC, soTmpQty, soTmpAmnt, soTmpDRls As String
soTmp = lblSO.Caption
soTmpDRls = DPDate_Released.Value

'Checks for insufficient stocks
For i = 1 To lvStockOut.ListItems.Count
    soTmpSC = lvStockOut.ListItems(i).Text
    soTmpQty = lvStockOut.ListItems(i).ListSubItems(2).Text
    soTmpAmnt = lvStockOut.ListItems(i).ListSubItems(3).Text
    
    'Outing the stock from Product table
    rs.Open "Select*from tblProduct where Product_ID='" & soTmpSC & "'", cn, 3, 3
    If Val(rs!Unit_In_Stock) < CInt(soTmpQty) Then
        ctr = ctr + 1
    End If
    Set rs = Nothing
Next

If ctr = 0 Then
    For i = 1 To lvStockOut.ListItems.Count
        soTmpSC = lvStockOut.ListItems(i).Text
        soTmpQty = lvStockOut.ListItems(i).ListSubItems(2).Text
        soTmpAmnt = lvStockOut.ListItems(i).ListSubItems(3).Text
        
        rs.Open "Select*from tblProduct where Product_ID='" & soTmpSC & "'", cn, 3, 3
        rs!Unit_In_Stock = Val(rs!Unit_In_Stock) - soTmpQty
        rs.Update
        Set rs = Nothing
        
        cn.Execute "Insert Into tblStockout(SO_No,Product_ID,Quantity,Amount,Date_Release)" & _
        "Values('" & soTmp & "','" & soTmpSC & "','" & soTmpQty & "','" & soTmpAmnt & "','" & soTmpDRls & "')"
    Next
    lvStockOut.ListItems.Clear
    cmdStockout.Enabled = False
    i_Clear.cLearMe Me
    i_Disable.Disable_Txt Me
    cmdAdd.Enabled = False
    soTmpSC = ""
    soTmpQty = ""
    soTmpAmnt = ""
    soTmpDRls = ""
    ctr = 0
    MsgBox "Stock out transaction No " & soTmp & " has been done.", vbInformation, "Inventory system"
    Call soLoadNum
Else
    MsgBox "Stock out transaction No " & soTmp & " has not been done due to" & vbCrLf & "some of the of the stocks are insufficient. Please check your" & vbCrLf & "quantity to be outed.", vbExclamation, "Inventory system"
End If
End Sub

Private Sub Form_Load()
DPDate_Released.Value = Now
Call soLoadNum

'Loads up the product ID
rs.Open "Select*from tblProduct", cn, 3, 3
cmbProdID.Clear
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        cmbProdID.AddItem rs!Product_ID
        rs.MoveNext
    Loop
End If
Set rs = Nothing

'Loads up the product names
rs.Open "Select*from tblProduct", cn, 3, 3
cmbProdName.Clear
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        cmbProdName.AddItem rs!Product_Name
        rs.MoveNext
    Loop
End If
Set rs = Nothing

End Sub




'auto numberer
Private Function soLoadNum()
rs.Open "select * from tblStockout Order By So_No DESC", cn, 3, 2
If rs.RecordCount = 0 Then
    lblSO.Caption = "SO-0000"
Else
    lblSO.Caption = "SO-" & Format(Right(rs!SO_NO, 4) + 1, "0000")
End If
rs.Close
End Function
