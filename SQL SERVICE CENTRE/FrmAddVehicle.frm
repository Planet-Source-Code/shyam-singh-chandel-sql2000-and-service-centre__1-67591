VERSION 5.00
Begin VB.Form FrmAddVehicle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Add Vehicle Type"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "FrmAddVehicle.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "Save"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Delete All"
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
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   1590
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   120
      Top             =   2760
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   120
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Add here a Vehicle Types"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Types"
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
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "FrmAddVehicle"
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
On Error Resume Next
Dim RSJO As New Recordset
 RSJO.Open "SELECT * FROM Vehicle", Connect, adOpenStatic, adLockOptimistic, adCmdText
      Do While Not RSJO.EOF
           RSJO.Delete
      RSJO.MoveNext
      Loop
      MsgBox "done"
      RSJO.Close
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim RSJO As New Recordset
 RSJO.Open "SELECT * FROM Vehicle where TypeofVehicle='" & List1.Text & "'", Connect, adOpenStatic, adLockOptimistic, adCmdText
       Do While Not RSJO.EOF
           RSJO.Delete
      RSJO.MoveNext
      Loop
       MsgBox List1.Text & "Deleted"
      RSJO.Close
      Form_Load
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim RSJO As New Recordset
 RSJO.Open "SELECT * FROM Vehicle", Connect, adOpenStatic, adLockOptimistic, adCmdText
      RSJO.AddNew
      RSJO!typeofVehicle = Text1.Text
      RSJO.Update
      MsgBox "Vehicle " & Text1.Text & " updated"
      RSJO.Close
     Form_Load
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
List1.Clear
Dim RSJO As New Recordset
 RSJO.Open "SELECT * FROM Vehicle", Connect, adOpenStatic, adLockOptimistic, adCmdText
      Do While Not RSJO.EOF
           List1.AddItem RSJO!typeofVehicle
      RSJO.MoveNext
      Loop
      RSJO.Close
End Sub


