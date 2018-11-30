VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmbill 
   Caption         =   "SELL PRODUCTS"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15765
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   15765
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtdiscount 
      Height          =   285
      Left            =   9000
      TabIndex        =   26
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "print"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   18
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "REMOVE FROM CART"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9600
      Picture         =   "frmbill.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD TO CART"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9600
      Picture         =   "frmbill.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   1455
   End
   Begin VB.ComboBox combopcompany 
      Height          =   315
      Left            =   7080
      TabIndex        =   13
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox comboptype 
      Height          =   315
      ItemData        =   "frmbill.frx":0F84
      Left            =   2640
      List            =   "frmbill.frx":0F8E
      TabIndex        =   12
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtbillmobno 
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtbillname 
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BILL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10920
      Picture         =   "frmbill.frx":0FAA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtbillqty 
      Height          =   615
      Left            =   11160
      TabIndex        =   4
      Top             =   5400
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   1215
      Left            =   600
      TabIndex        =   3
      Top             =   7320
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   6
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   1215
      Left            =   720
      TabIndex        =   2
      Top             =   5400
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape4 
      Height          =   3975
      Left            =   360
      Top             =   4680
      Width           =   12495
   End
   Begin VB.Shape Shape3 
      Height          =   3255
      Left            =   7080
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      Height          =   2775
      Left            =   360
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   6120
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "discount"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   25
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "net amount to be paid"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   24
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "gst(10%)"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "total"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   22
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lbltotal2 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   21
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lbltax 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   20
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lbltotal1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblbillid 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   17
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "enter quantity"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   14
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "select company"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "select category"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "mobile number"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "customer name"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "billid"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "BILL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdprint_Click()
DataEnvironment1.Connection1.Open
DataEnvironment1.Command1_Grouping 5001
DataReport1.Show
End Sub



Private Sub combopcompany_Click()
gridfill
End Sub


Private Sub comboptype_Click()
combopcompany.clear
If rs.State = 1 Then rs.Close
rs.Open "select distinct pcompany from product where ptype='" & comboptype.Text & "'", con, 3, 3
While Not rs.EOF
    combopcompany.AddItem (rs!pcompany)
rs.MoveNext
Wend
End Sub

Private Sub Command1_Click()
If rs.State = 1 Then rs.Close
rs.Open "bill", con, 3, 3
rs.addnew
rs!billid = Val(lblbillid.Caption)
rs!billname = txtbillname.Text
rs!billmobno = txtbillmobno.Text
'If rs1.State = 1 Then rs1.Close
'rs1.Open "select * from billtemp where billid=" & Val(lblbillid.Caption), con, 3, 3
'While Not rs1.EOF
'rs!billtotal = rs!billtotal + rs1!total
'rs1.MoveNext
'Wend
rs!billdate = Date
rs!billstot = lbltotal1.Caption
rs!billtax = Val(lbltax.Caption)
rs!billtotal = lbltotal2.Caption
rs.Update
MsgBox "bill calculated", , "Done"

DataEnvironment1.Connection1.Open
DataEnvironment1.Command1_Grouping lblbillid.Caption
DataReport1.Show

Unload Me
Me.Show
End Sub

Private Sub Command2_Click()

If rs.State = 1 Then rs.Close
rs.Open "select stqty from stock where pid=" & grid1.TextMatrix(grid1.RowSel, 1)
If Val(txtbillqty.Text) > rs!stqty Then
MsgBox "entered qty greater than available stock", , "Done"
txtbillqty.Text = ""
Else


billtempfill
grid2fill
con.Execute "update stock set stqty=stqty-" & Val(txtbillqty.Text) & "where pid=" & grid1.TextMatrix(grid1.RowSel, 1)
gridfill
txtbillqty.Text = ""
End If
End Sub

Private Sub Command3_Click()
con.Execute "update stock set stqty=stqty+" & grid2.TextMatrix(grid2.RowSel, 4) & "where pid=" & grid2.TextMatrix(grid2.RowSel, 1)
con.Execute "delete from billtemp where pid=" & grid2.TextMatrix(grid2.RowSel, 1)
grid2fill
gridfill
End Sub

Private Sub Form_Load()

If rs.State = 1 Then rs.Close
rs.Open "select * from billtemp", con, 3, 3
lblbillid.Caption = 5000
While Not rs.EOF
If Val(rs!billid) > Val(lblbillid.Caption) Then lblbillid.Caption = Val(rs!billid)
rs.MoveNext
Wend
lblbillid.Caption = Val(lblbillid.Caption) + 1



End Sub


Public Sub gridfill()
If rs.State = 1 Then rs.Close
'rs.Open "select * from stock where (pid in (select pid from product where pcompany='" & combopcompany.Text & "') )", con, 3, 3
rs.Open "select * from stock where (pid in (select pid from product where pcompany='" & combopcompany.Text & "' and ptype='" & comboptype.Text & "') )", con, 3, 3

grid1.Rows = 1
grid1.TextMatrix(0, 0) = "slno"
grid1.TextMatrix(0, 1) = "pid"
grid1.TextMatrix(0, 2) = "pmodelname"
grid1.TextMatrix(0, 3) = "stqty"
grid1.TextMatrix(0, 4) = "stprice"
Dim i As Integer
i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = i
grid1.TextMatrix(i, 1) = rs!PID
grid1.TextMatrix(i, 2) = rs!pmodelname
grid1.TextMatrix(i, 3) = rs!stqty
grid1.TextMatrix(i, 4) = rs!stprice
i = i + 1
rs.MoveNext
Wend
End Sub

Public Sub grid1_DblClick()
billtempfill
grid2fill
con.Execute "update stock set stqty=stqty-" & Val(txtbillqty.Text) & "where pid=" & grid1.TextMatrix(grid1.RowSel, 1)
gridfill
End Sub


Public Sub grid2fill()
lbltotal1.Caption = 0
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from billtemp where billid=" & Val(lblbillid.Caption), con, 3, 3
grid2.Rows = 1
grid2.TextMatrix(0, 0) = "slno"
grid2.TextMatrix(0, 1) = "pid"
grid2.TextMatrix(0, 2) = "pmodelname"
grid2.TextMatrix(0, 3) = "price"
grid2.TextMatrix(0, 4) = "qty"
grid2.TextMatrix(0, 5) = "total"
Dim j As Integer
j = 1
While Not rs1.EOF
grid2.Rows = grid1.Rows + 1
grid2.TextMatrix(j, 0) = j
grid2.TextMatrix(j, 1) = rs1!PID
grid2.TextMatrix(j, 2) = rs1!pmodelname
grid2.TextMatrix(j, 3) = rs1!price
grid2.TextMatrix(j, 4) = rs1!qty
grid2.TextMatrix(j, 5) = rs1!total
j = j + 1
lbltotal1.Caption = Val(lbltotal1.Caption) + rs1!total
rs1.MoveNext
Wend
End Sub

Public Sub billtempfill()

If rs1.State = 1 Then rs1.Close
rs1.Open "billtemp", con, 3, 3
rs1.addnew
'rs1!slno = Val(x)
rs1!PID = grid1.TextMatrix(grid1.RowSel, 1)
rs1!pmodelname = grid1.TextMatrix(grid1.RowSel, 2)
rs1!price = grid1.TextMatrix(grid1.RowSel, 4)
rs1!qty = Val(txtbillqty.Text)
rs1!total = (Val(rs1!price) * Val(rs1!qty))
rs1!billid = Val(lblbillid.Caption)
rs1.Update
End Sub


Public Sub grid2_DblClick()
con.Execute "update stock set stqty=stqty+" & grid2.TextMatrix(grid2.RowSel, 4) & "where pid=" & grid2.TextMatrix(grid2.RowSel, 1)
con.Execute "delete from billtemp where pid=" & grid2.TextMatrix(grid2.RowSel, 1)
grid2fill
gridfill
End Sub





Public Sub billslave()
End Sub

Private Sub lbltax_Change()
lbltotal2.Caption = 0
lbltotal2.Caption = Val(lbltotal2.Caption) + Val(lbltax.Caption) + Val(lbltotal1.Caption)
End Sub

Private Sub lbltotal1_Change()
lbltax.Caption = 0
lbltax.Caption = Val(lbltax.Caption) + Val(lbltotal1.Caption) * 0.1
End Sub

Private Sub txtdiscount_Change()
lbltotal2.Caption = 0
lbltotal2.Caption = Val(lbltotal2.Caption) + Val(lbltax.Caption) + Val(lbltotal1.Caption) - Val(txtdiscount.Text)
End Sub
