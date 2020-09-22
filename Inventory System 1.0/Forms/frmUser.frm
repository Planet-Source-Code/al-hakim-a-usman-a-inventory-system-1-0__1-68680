VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Profile (Administrator full access)"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      Height          =   975
      Left            =   -120
      TabIndex        =   11
      Top             =   -120
      Width           =   6495
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "USER PROFILE"
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
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   6375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "USER PROFILE"
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
         TabIndex        =   12
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   6135
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   615
         Left            =   5160
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   4320
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   615
         Left            =   3480
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   615
         Left            =   2640
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   615
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "Append"
         Enabled         =   0   'False
         Height          =   615
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6135
      Begin VB.TextBox txtPwd 
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
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "#"
         TabIndex        =   16
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtDepartment 
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
         TabIndex        =   4
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtDesignation 
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
         TabIndex        =   3
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtFulName 
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
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtUsrID 
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
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Password"
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
         TabIndex        =   17
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Department"
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
         TabIndex        =   8
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Full Name"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Designation"
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
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "User ID"
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
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid usrGrid 
      Height          =   1935
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000C&
      Height          =   2535
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      Height          =   3015
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   6135
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAppend_Click()
Dim usrID, uName, uDes, uDep, uPwd As String
'Checks for empty fields
Field_Check.Empty_Checks Me
If iTerminate = True Then
    iTerminate = False
    Exit Sub
End If

'Assigns the variable values
usrID = txtUsrID.Text
uName = txtFulName.Text
uDes = txtDesignation.Text
uDep = txtDepartment.Text
uPwd = txtPwd.Text
If ctrl_Flag = False Then
Call Redundancy_Check(usrID)
    If iTerminate = True Then
        MsgBox "The user you are trying to save is" & vbCrLf & "already on the record!", vbExclamation, "Inventory System"
        iTerminate = False
        Exit Sub
    End If
End If

'Saves in case of boolean expressions
Select Case ctrl_Flag
    Case False:
        rs.Open "Insert Into tblEmployee(Employee_ID,Employee_Name,Designation,Department,Password)" & _
        "Values('" & usrID & "','" & uName & "','" & uDes & "','" & uDep & "','" & uPwd & "')", cn, 3, 3
        Set rs = Nothing
        Call usrGrid_Load
        i_Clear.cLearMe Me
        i_Disable.Disable_Txt Me
        MsgBox "New user has been saved.", vbInformation, "Inventory System"
    Case True:
        rs.Open "Update tblEmployee set Employee_name='" & uName & "',Designation='" & uDes & "',Department='" & uDep & "',Password='" & uPwd & "'" & _
        "Where Employee_ID='" & iFind & "'", cn, 3, 3
        Set rs = Nothing
        Call usrGrid_Load
        i_Disable.Disable_Txt Me
        MsgBox "User account has been updated.", vbInformation, "Inventory System"
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
txtUsrID.Locked = False
cmdUpdate.Enabled = True
End Sub

Private Sub cmdDelete_Click()
iFind = InputBox("Enter user ID to delete.")
Call usr_Find
If ctr = 0 Then
    If iFind = "Administrator" Then
        i_Clear.cLearMe Me
        MsgBox "Deleting Administrator account is not allowed!", vbCritical, "Inventory System"
        Exit Sub
    Else
        rs.Open "Delete*from tblEmployee where Employee_ID='" & iFind & "'", cn, 3, 3
        Set rs = Nothing
        Call usrGrid_Load
        i_Clear.cLearMe Me
        MsgBox "User with User ID " & iFind & " has been deleted.", vbInformation, "Inventory System"
    End If
End If

End Sub

Private Sub cmdFind_Click()
iFind = InputBox("Enter user ID to find.")
Call usr_Find
End Sub

Private Sub cmdNew_Click()
i_Clear.cLearMe Me
i_Enable.Enable_Txt Me
ctrl_Flag = False
cmdAppend.Enabled = True
txtUsrID.SetFocus
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
iFind = InputBox("Enter user ID to update.")
Call usr_Find
If ctr = 0 Then
    ctrl_Flag = True
    cmdAppend.Enabled = True
    i_Enable.Enable_Txt Me
    txtUsrID.Locked = True
    txtFulName.SetFocus
    cmdUpdate.Enabled = False
End If

End Sub

Private Sub Form_Load()
Call usrGrid_Load

With usrGrid
    .ColWidth(0) = 200
    .ColWidth(1) = 1100
    .ColWidth(2) = 2000
End With
End Sub

Private Function usrGrid_Load()
rs.Open "Select*from tblEmployee", cn, 3, 3
Set usrGrid.DataSource = rs
Set rs = Nothing
End Function


'Redundancy checking function
Private Function Redundancy_Check(usrID)
rs.Open "Select*from tblEmployee where Employee_ID='" & usrID & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    If Not (rs.BOF And rs.EOF) Then
        iTerminate = True
    End If
End If
Set rs = Nothing
End Function




Private Sub usrGrid_Click()
X = usrGrid.Row
With usrGrid
    txtUsrID.Text = .TextMatrix(X, 1)
    txtFulName.Text = .TextMatrix(X, 2)
    txtDesignation.Text = .TextMatrix(X, 3)
    txtDepartment.Text = .TextMatrix(X, 4)
    txtPwd.Text = .TextMatrix(X, 5)
End With
End Sub
