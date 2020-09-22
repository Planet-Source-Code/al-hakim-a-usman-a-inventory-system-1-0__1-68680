VERSION 5.00
Begin VB.Form frmAdminAutorize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admin Authorization"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtAdminPwd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "#"
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdminAutorize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
ctr = 0
If rs.State = 1 Then Set rs = Nothing
rs.Open "Select*from tblEmployee", cn, 3, 3
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        If rs!Employee_ID = "Administrator" And rs!Password = txtAdminPwd.Text Then
            Set rs = Nothing
            Unload Me
            frmUser.Show
            ctr = 0
            Exit Do
        Else
            rs.MoveNext
            ctr = ctr + 1
        End If
    Loop
    If ctr > 0 Then
        Set rs = Nothing
        MsgBox "Incorrect pasword. You are not authorized" & vbCrLf & "to modify any user account!", vbCritical, "Inventory system"
        ctr = 0
    End If
End If
End Sub
