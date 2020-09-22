VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSupplier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier Profile"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid supGrid 
      Height          =   2295
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4048
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   7815
      Begin VB.TextBox txtSupID 
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
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtSupName 
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
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox txtAddress 
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
         TabIndex        =   6
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox txtTelephone 
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
         TabIndex        =   5
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Supplier ID"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Address"
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
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Supplier Name"
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
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telephone"
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
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   7815
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   615
         Left            =   6600
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   5520
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   615
         Left            =   4440
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   615
         Left            =   3360
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   615
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "Append"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      Caption         =   "Frame3"
      Height          =   1215
      Left            =   -240
      TabIndex        =   0
      Top             =   -240
      Width           =   8295
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER PROFILE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "SUPPLIER PROFILE"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6375
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      Height          =   3255
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   7815
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000C&
      Height          =   2175
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   7815
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAppend_Click()
Field_Check.Empty_Checks Me
If iTerminate = True Then
    iTerminate = False
    Exit Sub
End If


Sup_ID = UCase(txtSupID.Text)
Sup_Name = txtSupName.Text
sAdd = txtAddress.Text
sTel = txtTelephone.Text

If ctrl_Flag = False Then
    Call Redundancy_Check(Sup_ID)
    If iTerminate = True Then
        MsgBox "The supplier you are trying to save is" & vbCrLf & "already on the record!", vbExclamation, "Inventory System"
        iTerminate = False
        Exit Sub
    End If
End If


Select Case ctrl_Flag
    Case False:
        rs.Open "Insert Into tblSupplier(Supplier_ID,Supplier_Name,Address,Telephone)" & _
        "Values('" & Sup_ID & "','" & Sup_Name & "','" & sAdd & "','" & sTel & "')", cn, 3, 3
        Set rs = Nothing
        i_Clear.cLearMe Me
        i_Disable.Disable_Txt Me
        cmdAppend.Enabled = False
        Call grd_Data_Loader
        MsgBox "One new supplier has been saved.", vbInformation, "Inventory System"
    Case True:
        rs.Open "Update tblSupplier set Supplier_Name='" & Sup_Name & "',Address='" & sAdd & "',Telephone='" & sTel & "'" & _
        "Where Supplier_ID='" & Sup_ID & "'", cn, 3, 3
        i_Disable.Disable_Txt Me
        Set rs = Nothing
        Call grd_Data_Loader
        MsgBox "Supplier with supplier ID " & Sup_ID & " has been updated.", vbInformation, "Inventory System"
End Select
ctrl_Flag = False
cmdAppend.Enabled = False
cmdUpdate.Enabled = True
End Sub

Private Sub cmdCancel_Click()
ctrl_Flag = False
cmdAppend.Enabled = False
i_Clear.cLearMe Me
i_Disable.Disable_Txt Me
txtSupID.Locked = False
cmdUpdate.Enabled = True
End Sub

Private Sub cmdDelete_Click()
iFind = UCase(InputBox("Enter supplier ID to delete."))
Call sup_Find
If ctr = 0 Then
    rs.Open "Delete*from tblSupplier where Supplier_ID='" & UCase(iFind) & "'", cn, 3, 3
    Set rs = Nothing
    Call grd_Data_Loader
    i_Clear.cLearMe Me
    MsgBox "Supplier with supplier ID " & iFind & " has been deleted.", vbInformation, "Inventory System"
End If
End Sub

Private Sub cmdFind_Click()
iFind = UCase(InputBox("Enter supplier ID to find."))
Call sup_Find
End Sub

Private Sub cmdNew_Click()
i_Enable.Enable_Txt Me
i_Clear.cLearMe Me
txtSupID.SetFocus
cmdAppend.Enabled = True
txtSupID.Locked = False
ctrl_Flag = False
End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Function grd_Data_Loader()
rs.Open "Select*from tblSupplier", cn, 3, 3
Set supGrid.DataSource = rs
Set rs = Nothing
End Function

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
iFind = UCase(InputBox("Enter supplier ID to update."))
Call sup_Find
If ctr = 0 Then
    ctrl_Flag = True
    i_Enable.Enable_Txt Me
    txtSupID.Locked = True
    cmdAppend.Enabled = True
    cmdUpdate.Enabled = False
End If
End Sub

Private Sub Form_Load()
With supGrid
    .ColWidth(0) = 300
    .ColWidth(1) = 1500
    .ColWidth(2) = 3000
    .ColWidth(3) = 4000
    .ColWidth(4) = 1400
End With
Call grd_Data_Loader
End Sub



'Redundancy checking function
Private Function Redundancy_Check(Sup_ID)
rs.Open "Select*from tblSupplier where Supplier_ID='" & Sup_ID & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    If Not (rs.BOF And rs.EOF) Then
        iTerminate = True
    End If
End If
Set rs = Nothing
End Function




Private Sub supGrid_Click()
X = supGrid.Row
With supGrid
    txtSupID.Text = .TextMatrix(X, 1)
    txtSupName.Text = .TextMatrix(X, 2)
    txtAddress.Text = .TextMatrix(X, 3)
    txtTelephone.Text = .TextMatrix(X, 4)
End With









End Sub
