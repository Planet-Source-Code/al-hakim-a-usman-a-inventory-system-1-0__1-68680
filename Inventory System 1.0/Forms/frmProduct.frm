VERSION 5.00
Begin VB.Form frmProducts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      Caption         =   "Frame3"
      Height          =   1455
      Left            =   0
      TabIndex        =   25
      Top             =   -240
      Width           =   8775
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT DATABASE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   360
         TabIndex        =   27
         Top             =   405
         Width           =   7935
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "PRODUCT DATABASE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   7935
      End
   End
   Begin VB.ListBox lstProduct 
      Height          =   4350
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   3000
      TabIndex        =   12
      Top             =   4800
      Width           =   5295
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   615
         Left            =   4440
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   3720
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   615
         Left            =   3000
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   615
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   615
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "Append"
         Enabled         =   0   'False
         Height          =   615
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3000
      TabIndex        =   0
      Top             =   1320
      Width           =   5295
      Begin VB.ComboBox cboSupplier 
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
         Left            =   2040
         TabIndex        =   21
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtINStock 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox txtUnit_Price 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   2160
         Width           =   3135
      End
      Begin VB.ComboBox cboCategory 
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
         Left            =   2040
         TabIndex        =   9
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtProduct_Name 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtProduct_ID 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit_In_Stock"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit_Price"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Category"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Product_Name"
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
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Supplier"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Product_ID"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000C&
      Caption         =   "Label9"
      Height          =   4335
      Left            =   240
      TabIndex        =   24
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000C&
      Height          =   3015
      Left            =   3120
      TabIndex        =   22
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   3120
      TabIndex        =   20
      Top             =   4920
      Width           =   5295
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()

End Sub

Private Sub cmdAppend_Click()
Dim pIDCheck$
pIDCheck = txtProduct_ID

Field_Check.Empty_Checks Me
If iTerminate = True Then
    iTerminate = False
    Exit Sub
End If

If ctrl_Flag = False Then
    Call Redundancy_Check(pIDCheck)
    If iTerminate = True Then
        MsgBox "The product you are trying to save is" & vbCrLf & "already on the record!", vbExclamation, "Inventory System"
        iTerminate = False
        Exit Sub
    End If
End If

prod_ID = txtProduct_ID.Text
Call SQL_Execute(ctrl_Flag, prod_ID)
i_Clear.cLearMe Me
i_Disable.Disable_Txt Me
cmdAppend.Enabled = False
cmdUpdate.Enabled = True
ctrl_Flag = False
prod_ID = ""

End Sub


Private Sub cmdCancel_Click()
ctrl_Flag = False
i_Clear.cLearMe Me
i_Disable.Disable_Txt Me
cmdAppend.Enabled = False
cmdUpdate.Enabled = True
End Sub

Private Sub cmdDelete_Click()
i_Delete = InputBox("Enter Stock code of an item to delete.")
rs.Open "select*from tblProduct where Product_ID='" & i_Delete & "'", cn, 3, 3
rs.Delete
Set rs = Nothing
Call Reload_List
End Sub

Private Sub cmdFind_Click()
Dim pID As String
pID = InputBox("Please enter product ID to search")
Call Srch_Rec(Me.Name, pID)
End Sub

Private Sub cmdNew_Click()

cmdAppend.Enabled = True
ctrl_Flag = False
i_Enable.Enable_Txt Me
i_Clear.cLearMe Me
txtProduct_ID.SetFocus
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
ctrl_Flag = True
prod_ID = InputBox("Please enter product ID to update.")
If prod_ID = "" Then
    ctrl_Flag = False
    Exit Sub
Else
    Call Srch_Rec(Me.Name, prod_ID)
    i_Enable.Enable_Txt Me
    txtProduct_ID.Locked = True
    txtProduct_ID.SetFocus
    cmdAppend.Enabled = True
    cmdUpdate.Enabled = False
End If
End Sub

Private Sub Form_Load()
rs.Open "select*from tblProduct", cn, 3, 3
Do While Not rs.EOF
    lstProduct.AddItem rs!Product_Name
    rs.MoveNext
Loop
Set rs = Nothing

'Collects the Category from the table
rs.Open "Select distinct Category from tblProduct", cn, 3, 3
If rs.RecordCount > 0 Then
    Do While Not rs.EOF
        cboCategory.AddItem rs!category
        rs.MoveNext
    Loop
End If
Set rs = Nothing

rs.Open "Select*from tblSupplier", cn, 3, 3
If rs.RecordCount > 0 Then
    Do While Not rs.EOF
        cboSupplier.AddItem rs!supplier_name
        rs.MoveNext
    Loop
End If
Set rs = Nothing

End Sub


'Redundancy checking function
Private Function Redundancy_Check(pIDCheck)
rs.Open "Select*from tblProduct where Product_ID='" & pIDCheck & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    If Not (rs.BOF And rs.EOF) Then
        iTerminate = True
    End If
End If
Set rs = Nothing
End Function






Private Sub lstProduct_Click()
Dim pNam As String
If lstProduct.ListCount > 0 Then
    pNam = lstProduct.Text
    rs.Open "Select*from tblProduct where Product_Name='" & pNam & "'", cn, 3, 3
    txtProduct_ID.Text = rs!Product_ID
    txtProduct_Name.Text = rs!Product_Name
    cboSupplier.Text = rs!supplier
    cboCategory.Text = rs!category
    txtUnit_Price.Text = rs!Unit_Price
    txtINStock.Text = rs!Unit_In_Stock
    Set rs = Nothing
Else
     Exit Sub
End If
End Sub






