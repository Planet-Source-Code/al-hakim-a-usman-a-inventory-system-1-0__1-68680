VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Inventory System"
   ClientHeight    =   3645
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6870
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "Main.frx":0000
   Picture         =   "Main.frx":0442
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu lin 
         Caption         =   "Log In"
      End
      Begin VB.Menu lot 
         Caption         =   "Log Out"
         Enabled         =   0   'False
      End
      Begin VB.Menu ext 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu trans 
      Caption         =   "Transaction"
      Enabled         =   0   'False
      Begin VB.Menu sin 
         Caption         =   "Stock In"
      End
      Begin VB.Menu sot 
         Caption         =   "Stock Out"
      End
      Begin VB.Menu po 
         Caption         =   "Purchase Order"
      End
   End
   Begin VB.Menu maint 
      Caption         =   "Maintenance"
      Enabled         =   0   'False
      Begin VB.Menu user 
         Caption         =   "User"
      End
      Begin VB.Menu prod 
         Caption         =   "Product"
      End
      Begin VB.Menu sup 
         Caption         =   "Supplier"
      End
   End
   Begin VB.Menu rep 
      Caption         =   "Report"
      Enabled         =   0   'False
      Begin VB.Menu ap 
         Caption         =   "All Product"
      End
      Begin VB.Menu aps 
         Caption         =   "All Product by Supplier"
      End
      Begin VB.Menu apc 
         Caption         =   "All Product by Category"
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
Load frmAbout
End Sub

Private Sub ap_Click()
On Error Resume Next

Set rs = New ADODB.Recordset
rs.Open "SELECT * From tblProduct", cn
Set DataReport1.DataSource = rs.DataSource
For Each obj In DataReport1.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs.DataMember
    End If
Next
DataReport1.Sections("Section1").Controls("Text1").DataField = "Product_ID"
DataReport1.Sections("Section1").Controls("Text2").DataField = "Product_Name"
DataReport1.Sections("Section1").Controls("Text3").DataField = "Supplier"
DataReport1.Sections("Section1").Controls("Text4").DataField = "Category"
DataReport1.Sections("Section1").Controls("Text5").DataField = "Unit_Price"
DataReport1.Sections("Section1").Controls("Text6").DataField = "Unit_In_Stock"
DataReport1.Refresh
DataReport1.Show
Set rs = Nothing
End Sub



Private Sub apc_Click()
On Error Resume Next
Dim RPT$, RPT2$
RPT = InputBox("Enter Product Category.")

Set rs = New ADODB.Recordset
rs.Open "SELECT * From tblProduct where Category='" & RPT & "'", cn
RPT2 = rs!category
Set DataReport3.DataSource = rs.DataSource

For Each obj In DataReport3.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs.DataMember
    End If
Next
DataReport3.Sections("Section1").Controls("Text1").DataField = "Product_ID"
DataReport3.Sections("Section1").Controls("Text2").DataField = "Product_Name"
DataReport3.Sections("Section2").Controls("Label1").Caption = RPT2
DataReport3.Sections("Section1").Controls("Text3").DataField = "Supplier"
DataReport3.Sections("Section1").Controls("Text5").DataField = "Unit_Price"
DataReport3.Sections("Section1").Controls("Text6").DataField = "Unit_In_Stock"
DataReport3.Refresh
DataReport3.Show
Set rs = Nothing
End Sub

Private Sub aps_Click()
On Error Resume Next
Dim RPT$, RPT2$
RPT = InputBox("Enter product supplier name.")

Set rs = New ADODB.Recordset
rs.Open "SELECT * From tblProduct where Supplier='" & RPT & "'", cn
RPT2 = rs!supplier
Set DataReport2.DataSource = rs.DataSource

For Each obj In DataReport2.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs.DataMember
    End If
Next
DataReport2.Sections("Section1").Controls("Text1").DataField = "Product_ID"
DataReport2.Sections("Section1").Controls("Text2").DataField = "Product_Name"
DataReport2.Sections("Section2").Controls("Label1").Caption = RPT2
DataReport2.Sections("Section1").Controls("Text4").DataField = "Category"
DataReport2.Sections("Section1").Controls("Text5").DataField = "Unit_Price"
DataReport2.Sections("Section1").Controls("Text6").DataField = "Unit_In_Stock"
DataReport2.Refresh
DataReport2.Show
Set rs = Nothing
End Sub

Private Sub ext_Click()
End
End Sub

Private Sub lin_Click()
frmLogin.Show
End Sub

Private Sub lot_Click()
lin.Enabled = True
trans.Enabled = False
maint.Enabled = False
rep.Enabled = False
lot.Enabled = False
End Sub

Private Sub MDIForm_Load()
dBase = App.Path & "\Inventory.mdb"
cn.Open "Driver={Microsoft Access Driver (*.mdb)};dbq=" & dBase
End Sub

Private Sub po_Click()
frmPOrder.Show
End Sub

Private Sub prod_Click()
frmProducts.Show
End Sub

Private Sub sin_Click()
frmStockIn.Show
End Sub

Private Sub sot_Click()
frmStockout.Show
End Sub

Private Sub sup_Click()
frmSupplier.Show
End Sub

Private Sub user_Click()
frmAdminAutorize.Show
End Sub
