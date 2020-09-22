VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2325
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5955
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0442
   ScaleHeight     =   1373.688
   ScaleMode       =   0  'User
   ScaleWidth      =   5591.421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   7320
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   345
      Left            =   2730
      TabIndex        =   1
      Top             =   360
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2730
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   750
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   375
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'////                                                                    /////
'////     Developer: Shyam Singh Chandel                                 /////
'////     Jr. Technician (United News of India, Shillong)                /////
'////     URL http://tech.groups.yahoo.com/group/ssc_visual_basic/       /////
'////                                                                    /////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////


Option Explicit
Public LoginSucceeded As Boolean

Private Sub Command1_Click()
On Error Resume Next
    Dim RSS As New Recordset
    Dim U As String, P As String
    If txtUserName.Text = "" And txtPassword.Text = "" Then
    MsgBox "No User Name and Password is entered. Please Enter the fields."
    txtUserName.SetFocus
    Exit Sub
    End If
        RSS.Open "SELECT * FROM login where Password='" & txtPassword.Text & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
        U = RSS!user
        P = RSS!Password
        RSS.Close
        If txtUserName = U And txtPassword = P Then
     LoginSucceeded = True
     Me.Hide
     frm_main.Show
    Else
        MsgBox "Invalid User Name and Password, try again!", , "Login"
        txtPassword = ""
        txtUserName = ""
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Command2_Click()
     LoginSucceeded = False
    End
End Sub

Private Sub Form_Load()
  On Error Resume Next
    List1.Clear
    Dim RSS As New Recordset
        RSS.Open "SELECT * FROM login", Connect, adOpenStatic, adLockOptimistic, adCmdText
        Do While Not RSS.EOF
             List1.AddItem RSS!UserName
        RSS.MoveNext
      Loop
      RSS.Close
      
    If List1.ListCount <= 0 Then
    MsgBox "No User defind on the system. Please Enter the User Record First"
      FrmAddUsers.Show vbModal
     End If
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
LoginSucceeded = False
    End
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1_Click
End If

End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtPassword.SetFocus
End If

End Sub
