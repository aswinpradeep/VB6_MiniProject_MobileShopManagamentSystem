Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset

Public rsstock As New ADODB.Recordset
Public rspurchase As New ADODB.Recordset
Public rspurchasetemp As New ADODB.Recordset

Sub main()
'con.Open "miniproject"
'MDIForm1.Show
con.Open "miniproject"
'frmload.Show
frmlogin.Show
End Sub
