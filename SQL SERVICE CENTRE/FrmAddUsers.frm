VERSION 5.00
Begin VB.Form FrmAddUsers 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Users"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "FrmAddUsers.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Add User"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Delete Selected"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text3 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   1785
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   120
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      Height          =   2415
      Left            =   120
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
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
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAddUsers"
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

Private Sub Command1_Click()
Dim RSS As New Recordset
        RSS.Open "SELECT * FROM login where UserName='" & List1.Text & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
        RSS.Delete
        MsgBox "User " & "( " & List1.Text & " )" & " Deleted."
        RSS.Close
        Form_Load
End Sub

Private Sub Command2_Click()
Dim RSS As New Recordset
        RSS.Open "SELECT * FROM login where USER='" & txtUserName & "' and Password='" & txtPassword & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
              If RSS.EOF = True Then
                  RSS.AddNew
                      RSS!UserName = Text1
                      RSS!user = Text2
                      RSS!Password = Text3
                  RSS.Update
                  RSS.Close
                  MsgBox "User Added sucessfully"
                  blank
                  Form_Load
        Else
        MsgBox "USER " & Text1 & " already exist"
        RSS.Close
        blank
        Exit Sub
        End If
End Sub

Private Sub Command3_Click()
Unload Me
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
End Sub

Private Sub blank()
Text1 = ""
Text2 = ""
Text3 = ""

End Sub


Private Sub List1_Click()
        Dim RSS As New Recordset
        RSS.Open "SELECT * FROM login where UserName='" & List1.Text & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
        Text1 = RSS!UserName
        Text2 = RSS!user
        Text3 = RSS!Password
        RSS.Close
        
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text2.SetFocus
End If

End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text3.SetFocus
End If

End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Command2_Click
End If

End Sub


