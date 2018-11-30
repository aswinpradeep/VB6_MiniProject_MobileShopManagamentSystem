VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpurchase 
   Caption         =   "PURCHASE"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   15240
   Begin VB.TextBox txtdiscount 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12120
      TabIndex        =   31
      Top             =   6945
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   1215
      Left            =   720
      TabIndex        =   30
      Top             =   6360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   8
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "R"
      Height          =   255
      Left            =   6720
      TabIndex        =   27
      Top             =   2280
      Width           =   255
   End
   Begin VB.ComboBox comboptype 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmpurchase.frx":0000
      Left            =   3825
      List            =   "frmpurchase.frx":000A
      TabIndex        =   26
      Top             =   3600
      Width           =   2775
   End
   Begin VB.ComboBox combosid 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3825
      TabIndex        =   9
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txtpuqty 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11265
      TabIndex        =   8
      Top             =   3390
      Width           =   2775
   End
   Begin VB.TextBox txtpuprice 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11265
      TabIndex        =   7
      Top             =   2700
      Width           =   2775
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "add"
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
      Left            =   2760
      Picture         =   "frmpurchase.frx":0026
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Picture         =   "frmpurchase.frx":23F8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "DELETE"
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
      Left            =   5520
      Picture         =   "frmpurchase.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtputotal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11265
      TabIndex        =   2
      Top             =   4080
      Width           =   2775
   End
   Begin VB.ComboBox combopcompany 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3825
      TabIndex        =   1
      Top             =   4320
      Width           =   2775
   End
   Begin VB.ComboBox combopmodelname 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11265
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dtppudate 
      Height          =   375
      Left            =   3825
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   42979
   End
   Begin VB.Label lblx 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape3 
      Height          =   2655
      Left            =   9480
      Top             =   5160
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      Height          =   2655
      Left            =   480
      Top             =   5160
      Width           =   8535
   End
   Begin VB.Shape Shape1 
      Height          =   4335
      Left            =   360
      Top             =   600
      Width           =   13935
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "net amount to be paid"
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
      Left            =   2040
      TabIndex        =   35
      Top             =   8040
      Width           =   2775
   End
   Begin VB.Label Label15 
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
      Height          =   375
      Left            =   9840
      TabIndex        =   34
      Top             =   7005
      Width           =   1455
   End
   Begin VB.Label Label14 
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
      Height          =   375
      Left            =   9840
      TabIndex        =   33
      Top             =   6195
      Width           =   1455
   End
   Begin VB.Label Label13 
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
      Height          =   375
      Left            =   9840
      TabIndex        =   32
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblsname 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3825
      TabIndex        =   29
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "supplier name"
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
      Left            =   735
      TabIndex        =   28
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "product type"
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
      Left            =   735
      TabIndex        =   25
      Top             =   3570
      Width           =   2775
   End
   Begin VB.Label lbltotal2 
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
      Height          =   375
      Left            =   5280
      TabIndex        =   24
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label lbltax 
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
      Height          =   375
      Left            =   12120
      TabIndex        =   23
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label lbltotal1 
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
      Height          =   375
      Left            =   12120
      TabIndex        =   22
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblpid 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11265
      TabIndex        =   21
      Top             =   1995
      Width           =   2775
   End
   Begin VB.Label lblpuid 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3825
      TabIndex        =   20
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PURCHASE"
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
      Left            =   5160
      TabIndex        =   19
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "purchase id"
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
      Left            =   735
      TabIndex        =   18
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "supplier id"
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
      Left            =   735
      TabIndex        =   17
      Top             =   2235
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "product company"
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
      Left            =   735
      TabIndex        =   16
      Top             =   4215
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "quantity"
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
      Left            =   8175
      TabIndex        =   15
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "buying price"
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
      Left            =   8175
      TabIndex        =   14
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "pudate"
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
      Left            =   735
      TabIndex        =   13
      Top             =   1545
      Width           =   2775
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
      Height          =   375
      Left            =   8175
      TabIndex        =   12
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "model name"
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
      Left            =   8175
      TabIndex        =   11
      Top             =   1290
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "pid"
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
      Left            =   8175
      TabIndex        =   10
      Top             =   1920
      Width           =   2775
   End
End
Attribute VB_Name = "frmpurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdreset_Click()
MsgBox "make sure you have deleted all items manually before proceeding", vbOKOnly, "Warning"
combosid.Enabled = True
lblsname.Caption = ""
End Sub

Private Sub cmdsave_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from purchasetemp where puid=" & Val(lblpuid.Caption), con, 3, 3
'gtotal = 0
'While Not rs.EOF
'gtotal = gtotal + Val(rs!putotal)
'rs.MoveNext
'Wend
'gridfill
If rs1.State = 1 Then rs1.Close
rs1.Open "purchase", con, 3, 3
rs1.addnew
rs1!puid = Val(lblpuid.Caption)
rs1!sid = combosid.Text
rs1!pudate = dtppudate.Value
rs1!pugtotal = Val(lbltotal2.Caption)
rs1.Update
MsgBox "purchase done", vbOKOnly, "Done"
Unload Me
Me.Show
'Me.Hide
'Me.Show
'stockfill
'fnfinal
End Sub



Private Sub Combo1_Change()

End Sub

Private Sub Combo1_Click()

End Sub

Private Sub combopcompany_Click()
combopmodelname.clear
If rs.State = 1 Then rs.Close
rs.Open "select distinct pmodelname from product where pcompany='" & combopcompany.Text & "' and ptype='" & comboptype.Text & "'", con, 3, 3
While Not rs.EOF
combopmodelname.AddItem (rs!pmodelname)
rs.MoveNext
Wend
End Sub

Private Sub combopmodelname_Change()



'If rs.State = 1 Then rs.Close
'rs.Open "select pid from product where pcompany='" & combopcompany.Text & "' and pmodelname='" & combopmodelname.Text & "'", con, 3, 3

'combopid.Text = rs!PID
End Sub

Private Sub combopmodelname_Click()
If rs.State = 1 Then rs.Close
rs.Open "select pid from product where pcompany='" & combopcompany.Text & "' and pmodelname='" & combopmodelname.Text & "'", con, 3, 3

lblpid.Caption = rs!PID
End Sub

Private Sub Command1_Click()
modelnamefill
End Sub

Private Sub comboptype_Click()
combosid.Enabled = False
combopcompany.clear
If rs.State = 1 Then rs.Close
rs.Open "select distinct pcompany from product where ptype='" & comboptype.Text & "'", con, 3, 3
While Not rs.EOF
combopcompany.AddItem (rs!pcompany)
rs.MoveNext
Wend

End Sub

Private Sub combosid_Click()
If rs.State = 1 Then rs.Close
rs.Open "select sname from supplier where sid=" & combosid.Text, con, 3, 3
lblsname.Caption = rs!sname
End Sub

Private Sub Form_Load()
dtppudate.Value = Date
If rs.State = 1 Then rs.Close
rs.Open "select * from purchasetemp", con, 3, 3
lblpuid.Caption = 100
While Not rs.EOF
If Val(rs!puid) > Val(lblpuid.Caption) Then lblpuid.Caption = Val(rs!puid)
rs.MoveNext
Wend
lblpuid.Caption = Val(lblpuid.Caption) + 1


If rs.State = 1 Then rs.Close
rs.Open "select distinct sid from supplier", con, 3, 3
While Not rs.EOF
combosid.AddItem (rs!sid)
rs.MoveNext
Wend

End Sub
Public Sub gridfill()
lbltotal1.Caption = 0
If rs.State = 1 Then rs.Close
rs.Open "select * from purchasetemp where puid=" & Val(lblpuid.Caption), con, 3, 3
grid1.Rows = 1
grid1.TextMatrix(0, 0) = "puid"
grid1.TextMatrix(0, 1) = "pudate"
grid1.TextMatrix(0, 2) = "sid"
grid1.TextMatrix(0, 3) = "pid"
grid1.TextMatrix(0, 4) = "puprice"
grid1.TextMatrix(0, 5) = "puqty"
grid1.TextMatrix(0, 6) = "putotal"


Dim i As Integer
i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = rs!puid
grid1.TextMatrix(i, 1) = rs!pudate
grid1.TextMatrix(i, 2) = rs!sid
grid1.TextMatrix(i, 3) = rs!PID
grid1.TextMatrix(i, 4) = rs!puprice
grid1.TextMatrix(i, 5) = rs!puqty
grid1.TextMatrix(i, 6) = rs!putotal

lbltotal1.Caption = Val(lbltotal1.Caption) + rs!putotal
i = i + 1
rs.MoveNext
Wend
End Sub


Private Sub cmdadd_Click()
If rs.State = 1 Then rs.Close
rs.Open "purchasetemp", con, 3, 3

rs.addnew
rs!puid = Val(lblpuid.Caption)
rs!sid = combosid.Text
rs!PID = lblpid.Caption
rs!puprice = txtpuprice.Text
rs!puqty = txtpuqty.Text
rs!pudate = dtppudate.Value
rs!putotal = txtputotal.Text
rs.Update
rs.Close

gridfill
STOCKFILLX
clear
End Sub

Public Sub clear()
'txtpuid.Text = ""
comboptype.Text = ""
combopcompany.Text = ""
combopmodelname.Text = ""
lblpid.Caption = ""
txtpuprice.Text = ""
txtpuqty.Text = ""
txtputotal.Text = ""
dtppudate.Value = Date
End Sub
Private Sub cmddel_Click()

con.Execute "update stock set stqty=stqty+" & grid1.TextMatrix(grid1.RowSel, 5) & "where pid=" & grid1.TextMatrix(grid1.RowSel, 3)
con.Execute "delete from purchasetemp where pid=" & grid1.TextMatrix(grid1.RowSel, 3)



rs.Close
rs.Open "purchasetemp", con, 3, 3

gridfill
End Sub




Public Sub STOCKFILLX()
Dim count As Integer
count = 0
If rs1.State = 1 Then rs1.Close
rs1.Open "stock", con, 3, 3
While Not rs1.EOF
If rs1!PID = Val(lblpid.Caption) Then
    rs1!stqty = rs1!stqty + Val(txtpuqty.Text)
    count = 1
End If
rs1.MoveNext
Wend
If count = 0 Then
addnew
End If

End Sub
Public Sub addnew()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from stock", con, 3, 3

If rs.State = 1 Then rs.Close
rs.Open "select pmrp from product where pid=" & Val(lblpid.Caption), con, 3, 3

'rs1.addnew
'rs1!stid = (Val(lblpuid.Caption) * 10)

lblx.Caption = 1750
While Not rs1.EOF
If Val(rs1!stid) > Val(lblx.Caption) Then lblx.Caption = Val(rs1!stid)
rs1.MoveNext
Wend
lblx.Caption = Val(lblx.Caption) + 1

rs1.addnew
'rs1!stid = (Val(lblpuid.Caption) * 10)
rs1!stid = Val(lblx.Caption)
lblpid.Caption = Val(lblpid.Caption)
rs1!PID = Val(lblpid.Caption)
rs1!pmodelname = combopmodelname.Text
rs1!stprice = rs!pmrp
rs1!stqty = Val(txtpuqty.Text)
rs1.Update
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

Private Sub txtpuprice_Change()

txtputotal.Text = Val(txtpuprice.Text) * Val(txtpuqty.Text)

End Sub

Private Sub txtpuqty_Change()
txtputotal.Text = Val(txtpuprice.Text) * Val(txtpuqty.Text)
End Sub

