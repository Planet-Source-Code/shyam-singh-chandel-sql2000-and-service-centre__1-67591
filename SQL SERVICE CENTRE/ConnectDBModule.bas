Attribute VB_Name = "ConnectDBModule"
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
Global Connect As Connection
Global serverName As String
Global SR As Integer, SRPLUS As Integer, NR As Integer, NRPLUS As Integer
Global BillStat As String
Global DBPATH As String
    
Public Sub con()
On Error GoTo Handler
Set Connect = New Connection
Connect.Open "Provider=SQLOLEDB.1;Persist Security Info=True;Password=" & FrmServer.Text3.Text & ";User ID=" & FrmServer.Text2.Text & ";Initial Catalog=ServiceCentre;Data Source=" & FrmServer.Text1.Text
Exit Sub
Handler:
    MsgBox "Server Technical Problems!!" & Chr(13) & Chr(13) & _
    "Eighter Record has been deleted of not exist. !!" & Chr(13) & Chr(13) & _
    "Plese contect to US SOFTWARES for further information !!" & Chr(13) & Chr(13) & _
    "Error Description : " & Err.Description & Chr(13) & Chr(13) & _
    "Email: ussoftwares@rediffmail.com or shyamschandel@rediffmail.com", vbCritical
    End
End Sub

