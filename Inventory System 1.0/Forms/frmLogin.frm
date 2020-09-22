VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log In"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Log In"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtPwd 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtUsr 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
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
      TabIndex        =   3
      Top             =   720
      Width           =   1455
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
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdLogin_Click()
Dim uSr, pWd As String
uSr = txtUsr.Text
pWd = txtPwd.Text

rs.Open "Select*From tblEmployee", cn, 3, 3
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        If uSr = rs!Employee_ID And pWd = rs!Password Then
            Set rs = Nothing
                With MDIForm1
                    .lot.Enabled = True
                    .trans.Enabled = True
                    .maint.Enabled = True
                    .rep.Enabled = True
                    .lin.Enabled = False
                End With
                ctr = 0
                Unload Me
                Exit Do
        Else
            rs.MoveNext
            ctr = ctr + 1
        End If
    Loop
    If ctr > 0 Then
        Set rs = Nothing
        MsgBox "Access Denied!!", vbCritical, "Inventory System"
        i_Clear.cLearMe Me
        txtUsr.SetFocus
        ctr = 0
    End If
End If
End Sub
