VERSION 5.00
Begin VB.Form FrmSettings 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "FrmSettings.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   1555
      Left            =   240
      ScaleHeight     =   1530
      ScaleWidth      =   3060
      TabIndex        =   2
      Top             =   240
      Width           =   3085
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "FrmSettings.frx":0442
         Top             =   120
         Width           =   2595
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "Regenerate Job Card and Bill No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Delete Job Card and Bill No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "FrmSettings"
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

Private Sub Command2_Click()
Dim RSc As New Recordset
RSc.Open "SELECT * FROM NUMBERS", Connect, adOpenStatic, adLockOptimistic, adCmdText
       Do While Not RSc.EOF
        RSc.Delete
        RSc.MoveNext
        Loop
        RSc.Close
        MsgBox "done"
End Sub

Private Sub Command3_Click()
Dim RSc As New Recordset
RSc.Open "SELECT * FROM NUMBERS", Connect, adOpenStatic, adLockOptimistic, adCmdText
        RSc.AddNew
              RSc!BillNo = 0
              RSc!JobCard = 0
              RSc!RecNo = 0
        RSc.Update
        RSc.Close
        MsgBox "done"
End Sub

