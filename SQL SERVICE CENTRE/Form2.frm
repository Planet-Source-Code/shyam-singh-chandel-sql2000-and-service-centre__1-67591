VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Welcome to US SOFTWARES"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   1160
      Left            =   240
      ScaleHeight     =   1125
      ScaleWidth      =   3060
      TabIndex        =   7
      Top             =   240
      Width           =   3085
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "Form2.frx":0442
         Top             =   120
         Width           =   2595
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   7920
      Top             =   1080
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Network Application (SQL Server 2000 in Back end)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   360
      Width           =   4335
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   495
      Left            =   240
      Top             =   1500
      Width           =   7575
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   1335
      Left            =   240
      Top             =   2040
      Width           =   7575
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   1215
      Left            =   3435
      Top             =   240
      Width           =   4395
   End
   Begin VB.Shape Shape1 
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SOFTWARES "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " SERVICING CENTRE  1.2 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   645
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:-  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "shyamschandel@rediffmail.com     "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1590
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS OF THE SOFTWARE COMPANY SHOULD BE HERE."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD HERE A COMPANY NAME"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   7575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   1190
      Left            =   240
      Top             =   240
      Width           =   3120
   End
End
Attribute VB_Name = "Form2"
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


Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then
Timer1.Enabled = False
MsgBox "This Application is already running."
End
End If
End Sub

Private Sub Timer1_Timer()
frmLogin.Show
Unload Me
End Sub
