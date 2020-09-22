VERSION 5.00
Begin VB.Form frm_main 
   ClientHeight    =   5355
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10425
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnu_users 
      Caption         =   "Users"
   End
   Begin VB.Menu mnu_logoff 
      Caption         =   "Log Off"
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
updatecaption
End Sub

Private Sub mnu_logoff_Click()
Unload frm_useredit         '
Unload frm_users

Me.Hide
frm_passwordscreen.txt_user.Text = ""
frm_passwordscreen.txt_Password.Text = ""
frm_passwordscreen.Show 1
End Sub

Private Sub mnu_users_Click()
frm_users.Show
End Sub
Public Sub updatecaption()
Me.Caption = loggedin
End Sub
