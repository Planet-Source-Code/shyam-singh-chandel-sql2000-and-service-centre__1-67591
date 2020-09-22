VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Servicing Center Software"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   11775
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   1160
      Left            =   2760
      ScaleHeight     =   1125
      ScaleWidth      =   8700
      TabIndex        =   16
      Top             =   7080
      Width           =   8730
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "Form3.frx":0442
         Stretch         =   -1  'True
         Top             =   120
         Width           =   8235
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   0
      ScaleHeight     =   8505
      ScaleWidth      =   2385
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command10 
         BackColor       =   &H00808080&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7320
         Width           =   2055
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00808080&
         Caption         =   "Add Type of Vehicle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00808080&
         Caption         =   "Creat User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5160
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Servicing Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "Billing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00808080&
         Caption         =   "Unpaid Bills"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00808080&
         Caption         =   "Paid Bills"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00808080&
         Caption         =   "Reset Bill No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00808080&
         Caption         =   "Reset Job Card No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00808080&
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   $"Form3.frx":8684
         Height          =   255
         Left            =   11640
         TabIndex        =   13
         Top             =   2760
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
      Begin MSComctlLib.ImageList IList 
         Left            =   3600
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":8725
               Key             =   "LOGIN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":93FF
               Key             =   "ADDUSER"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":A0D9
               Key             =   "MEMO"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":ADB3
               Key             =   "SAVE"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":BA8D
               Key             =   "BACKUP"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":C767
               Key             =   "ITEMS"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":D441
               Key             =   "ACCNTLIST"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":E11B
               Key             =   "APP"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":EF6D
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":FDBF
               Key             =   "SERIALS"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":10699
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":10F73
               Key             =   "PRINT"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":1128D
               Key             =   "REPORTS"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":11F67
               Key             =   "TRASH"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":12C41
               Key             =   "DEPARTMENT"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":1351B
               Key             =   "USERS"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form3.frx":13DF5
               Key             =   "CONSOLE"
            EndProperty
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   1215
         Left            =   1080
         Picture         =   "Form3.frx":146CF
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   1215
         Left            =   1680
         Picture         =   "Form3.frx":1AE91
         Stretch         =   -1  'True
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Image5 
         Height          =   1215
         Left            =   3000
         Picture         =   "Form3.frx":21653
         Stretch         =   -1  'True
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   4320
         Picture         =   "Form3.frx":27E15
         Top             =   2760
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image7 
         Height          =   1215
         Left            =   5280
         Picture         =   "Form3.frx":28244
         Stretch         =   -1  'True
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Image8 
         Height          =   1215
         Left            =   6720
         Picture         =   "Form3.frx":28686
         Stretch         =   -1  'True
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Image9 
         Height          =   1215
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image BottonRepair 
         Height          =   480
         Left            =   4320
         Picture         =   "Form3.frx":28AC8
         ToolTipText     =   "Database Diagnostics"
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image BottonBackUp 
         Height          =   480
         Left            =   4320
         Picture         =   "Form3.frx":29792
         ToolTipText     =   "Backup Data"
         Top             =   3840
         Width           =   480
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Servicing Center Software"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   765
      Left            =   3030
      TabIndex        =   15
      Top             =   195
      Width           =   8010
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Servicing Center Software"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1000
      Left            =   3120
      TabIndex        =   14
      Top             =   240
      Width           =   8010
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1215
      Left            =   4920
      TabIndex        =   1
      Top             =   3120
      Width           =   5775
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   8535
      Left            =   120
      Top             =   0
      Width           =   11655
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu MnuSCE 
         Caption         =   "Servicing Center Entry"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu Billing 
         Caption         =   "Billing"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu paidbills 
         Caption         =   "Paid Bills"
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu unpbills 
         Caption         =   "Unpaid Bills"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frm_main"
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
Form1.Show
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Image7.Picture
    Label2.Caption = "Allow to enter the record of the vehicle, which comes for servicing."
    Command1.BackColor = vbRed
    Command10.BackColor = &H808080
    Command2.BackColor = &H808080
    Command3.BackColor = &H808080
    Command4.BackColor = &H808080
    Command5.BackColor = &H808080
    Command6.BackColor = &H808080
    Command7.BackColor = &H808080
    Command8.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub

Private Sub Command10_Click()
Dialog.Show
Unload Me

End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = IList.ListImages(9).Picture
    Label2.Caption = "Allow to Close the Software."
    Command10.BackColor = vbRed
    Command1.BackColor = &H808080
    Command2.BackColor = &H808080
    Command3.BackColor = &H808080
    Command4.BackColor = &H808080
    Command5.BackColor = &H808080
    Command6.BackColor = &H808080
    Command7.BackColor = &H808080
    Command8.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub

Private Sub Command2_Click()
Form4.Show

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Image3.Picture
    Label2.Caption = "Allow to View and Print bills and Job Card of the servicing vehicle and allow to move the bill to the 'Paid' or 'Unpaid' bill Section."
    Command2.BackColor = vbRed
    Command1.BackColor = &H808080
    Command10.BackColor = &H808080
    Command3.BackColor = &H808080
    Command4.BackColor = &H808080
    Command5.BackColor = &H808080
    Command6.BackColor = &H808080
    Command7.BackColor = &H808080
    Command8.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub

Private Sub Command3_Click()
Form6.Show

End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Image5.Picture
    Label2.Caption = "Allow to View and  Print 'Unpaid' bills and Job Card of the servicing vehicle and allow to move the bill to the 'Paid' bill Section."
    Command3.BackColor = vbRed
    Command1.BackColor = &H808080
    Command2.BackColor = &H808080
    Command10.BackColor = &H808080
    Command4.BackColor = &H808080
    Command5.BackColor = &H808080
    Command6.BackColor = &H808080
    Command7.BackColor = &H808080
    Command8.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub

Private Sub Command4_Click()
Form5.Show
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Image4.Picture
    Label2.Caption = "Allow to View and Print 'Paid' bills of the servicing vehicle and also allow to move the bill into 'Unpaid' bill Section. All the paid bill can find here."
    Command4.BackColor = vbRed
    Command1.BackColor = &H808080
    Command2.BackColor = &H808080
    Command3.BackColor = &H808080
    Command10.BackColor = &H808080
    Command5.BackColor = &H808080
    Command6.BackColor = &H808080
    Command7.BackColor = &H808080
    Command8.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub


Private Sub Command5_Click()
Dim INA As String, SQL As String

INA = InputBox("Enter bill no for reset")
If INA = "" Then
Exit Sub
Else
SQL = "SELECT BILLNO FROM NUMBERS"
Dim RSc As New Recordset
 RSc.Open SQL, Connect, adOpenStatic, adLockOptimistic, adCmdText
         currentMode = EditMode
             RSc!BillNo = INA
        RSc.Update
  RSc.Close
  MsgBox "RESET BILL NO DONE."
  End If

End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = IList.ListImages(6).Picture
    Label2.Caption = "Allow to Reset the Number of the Auto Bill No, which can set the Bill Numbers of eighter begining of middle."
    Command5.BackColor = vbRed
    Command1.BackColor = &H808080
    Command2.BackColor = &H808080
    Command3.BackColor = &H808080
    Command4.BackColor = &H808080
    Command10.BackColor = &H808080
    Command6.BackColor = &H808080
    Command7.BackColor = &H808080
    Command8.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub

Private Sub Command6_Click()
Dim INA As String
Dim SQL As String
INA = InputBox("Enter Job Card no for reset")
If INA = "" Then
Exit Sub
Else
SQL = "SELECT JobCard FROM NUMBERS"
Dim RSc As New Recordset
 RSc.Open SQL, Connect, adOpenStatic, adLockOptimistic, adCmdText
        currentMode = EditMode
             RSc!JobCard = INA
        RSc.Update
 RSc.Close
MsgBox "Reset of Job Card is Done"
End If

End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = IList.ListImages(6).Picture
    Label2.Caption = "Allow to Reset the Number of the Auto Job Card, which can set the Job Card Numbers of eighter begining of middle."
    Command6.BackColor = vbRed
    Command1.BackColor = &H808080
    Command2.BackColor = &H808080
    Command3.BackColor = &H808080
    Command4.BackColor = &H808080
    Command5.BackColor = &H808080
    Command10.BackColor = &H808080
    Command7.BackColor = &H808080
    Command8.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub

Private Sub Command7_Click()
FrmSettings.Show
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = BottonRepair.Picture
    Label2.Caption = "Allow to Reset the Number of the Auto Job Card and Bill No. if the Automatically Job Card and Bill No is not working."
    Command7.BackColor = vbRed
    Command1.BackColor = &H808080
    Command2.BackColor = &H808080
    Command3.BackColor = &H808080
    Command4.BackColor = &H808080
    Command5.BackColor = &H808080
    Command6.BackColor = &H808080
    Command10.BackColor = &H808080
    Command8.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub

Private Sub Command8_Click()
FrmAddUsers.Show
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Image6.Picture
    Label2.Caption = "Allow to Creat the Users for the Service Center Software."
    Command8.BackColor = vbRed
    Command1.BackColor = &H808080
    Command2.BackColor = &H808080
    Command3.BackColor = &H808080
    Command4.BackColor = &H808080
    Command5.BackColor = &H808080
    Command6.BackColor = &H808080
    Command7.BackColor = &H808080
    Command10.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub

Private Sub Command9_Click()
FrmAddVehicle.Show

End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Image8.Picture
    Label2.Caption = "Allow to Add the Vehicle Names."
    Command9.BackColor = vbRed
    Command1.BackColor = &H808080
    Command2.BackColor = &H808080
    Command3.BackColor = &H808080
    Command4.BackColor = &H808080
    Command5.BackColor = &H808080
    Command6.BackColor = &H808080
    Command7.BackColor = &H808080
    Command8.BackColor = &H808080
    Command10.BackColor = &H808080
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = Image9.Picture
    Label2.Caption = ""
    Command1.BackColor = &H808080
    Command10.BackColor = &H808080
    Command2.BackColor = &H808080
    Command3.BackColor = &H808080
    Command4.BackColor = &H808080
    Command5.BackColor = &H808080
    Command6.BackColor = &H808080
    Command7.BackColor = &H808080
    Command8.BackColor = &H808080
    Command9.BackColor = &H808080
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dialog.Show
Unload Me
End Sub

