VERSION 5.00
Begin VB.Form FrmServer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Connection Database"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "FrmServer.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   3240
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Text            =   "ServiceCentre"
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Cancel"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Connect"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   120
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   360
      Width           =   1335
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
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "FrmServer"
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
        USDB = Text4
        USPass = Text3
        USUser = Text2
        USServer = Text1
 con
        Form2.Show
        Call SaveSetting("USSQL", "Server", "ServerName", Text1.Text)
        Call SaveSetting("USSQL", "Server", "UserName", Text2.Text)
        Call SaveSetting("USSQL", "Server", "Password", Text3.Text)
        Call SaveSetting("USSQL", "Server", "Database", Text4.Text)
        Unload Me

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Text1.Text = GetSetting("USSQL", "Server", "ServerName", Text1.Text)
Text2.Text = GetSetting("USSQL", "Server", "UserName", Text2.Text)
Text3.Text = GetSetting("USSQL", "Server", "Password", Text3.Text)
Text4.Text = GetSetting("USSQL", "Server", "Database", Text4.Text)
If Text1.Text = "" Then
Timer1.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
        con
        Form2.Show
        Unload Me
        Timer1.Enabled = False
        
End Sub
