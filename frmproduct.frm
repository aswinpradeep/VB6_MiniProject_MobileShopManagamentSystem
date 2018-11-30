VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmproduct 
   Caption         =   "PRODUCTS"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15840
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   15840
   Begin VB.ComboBox comboptype 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmproduct.frx":0000
      Left            =   4080
      List            =   "frmproduct.frx":000A
      TabIndex        =   16
      Top             =   2760
      Width           =   3135
   End
   Begin VB.ComboBox combosearch 
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
      Left            =   13200
      TabIndex        =   13
      Text            =   "search"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14640
      TabIndex        =   12
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "edit"
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
      Left            =   11400
      Picture         =   "frmproduct.frx":0026
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "delete"
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
      Picture         =   "frmproduct.frx":0C68
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtpmrp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   9
      Top             =   5760
      Width           =   3135
   End
   Begin VB.ComboBox combopcompany 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4080
      TabIndex        =   7
      Top             =   3660
      Width           =   3135
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
      Left            =   7920
      Picture         =   "frmproduct.frx":18AA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   2055
      Left            =   7680
      TabIndex        =   3
      Top             =   1680
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtpmodelname 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Shape Shape4 
      Height          =   495
      Left            =   5520
      Top             =   240
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   13080
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      Height          =   1215
      Left            =   7800
      Top             =   3960
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      Height          =   5175
      Left            =   360
      Top             =   1320
      Width           =   7095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "ptype"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   15
      Top             =   2610
      Width           =   2655
   End
   Begin VB.Label lblpid 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   14
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "MRP"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "company"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3660
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "PRODUCT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "pname"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   4710
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "pid"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
End
Attribute VB_Name = "frmproduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddelete_Click()
x = MsgBox("Product deletion is a complex and will remove the traces of the same product from all included tables.Hence note down the detaile before you proceed.continue?", vbOKCancel, "Warning")
If x = vbCancel Then
Me.Show
Else
con.Execute "delete from billtemp where pid=" & Val(lblpid.Caption)
con.Execute "delete from purchasetemp where pid=" & Val(lblpid.Caption)
con.Execute "delete from replacement where pid=" & Val(lblpid.Caption)
con.Execute "delete from returntemp where pid=" & Val(lblpid.Caption)
con.Execute "delete from stock where pid=" & Val(lblpid.Caption)
con.Execute "delete from product where pid=" & Val(lblpid.Caption)
'con.Execute "delete from product where pid=" & Val(lblpid.Caption)
rs.Close
rs.Open "select * from product", con, 3, 3
clear
gridfill

End If
End Sub



Private Sub cmdreset_Click()
clear
disableall
cmdadd.Caption = "add"
cmdedit.Caption = "edit"
combosearch.Text = "Search"
gridfill
End Sub

Private Sub cmdedit_Click()
If cmdedit.Caption = "edit" Then
'clear
enableall
cmdedit.Caption = "save"
Else
'rs.addnew
If rs.State = 1 Then rs.Close
rs.Open "select * from product where pid=" & Val(lblpid.Caption), con, 3, 3
rs!PID = Val(lblpid.Caption)
rs!pcompany = combopcompany.Text
rs!ptype = comboptype.Text
rs!pmodelname = txtpmodelname.Text
rs!pmrp = txtpmrp.Text
rs.Update
cmdedit.Caption = "edit"
rs.Close
rs.Open "select * from product", con, 3, 3
gridfill
clear
disableall
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub combosearch_Change()
If rs.State = 1 Then rs.Close
rs.Open "select * from product where pcompany like'" & combosearch.Text & "%'", con, 3, 3
grid1.Rows = 1
grid1.TextMatrix(0, 0) = "pid"
grid1.TextMatrix(0, 1) = "ptype"
grid1.TextMatrix(0, 2) = "pcompany"

grid1.TextMatrix(0, 3) = "pmodelname"
grid1.TextMatrix(0, 4) = "pmrp"

Dim i As Integer
i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = rs!PID
grid1.TextMatrix(i, 1) = rs!ptype
grid1.TextMatrix(i, 2) = rs!pcompany

grid1.TextMatrix(i, 3) = rs!pmodelname
grid1.TextMatrix(i, 4) = rs!pmrp

i = i + 1
rs.MoveNext
Wend
End Sub

Private Sub Form_Load()
disableall
If rs.State = 1 Then rs.Close
rs.Open "select distinct pcompany from product", con, 3, 3
While Not rs.EOF
combopcompany.AddItem (rs!pcompany)
rs.MoveNext
Wend
gridfill
End Sub
Public Sub gridfill()
If rs.State = 1 Then rs.Close
rs.Open "select * from product", con, 3, 3
grid1.Rows = 1
grid1.TextMatrix(0, 0) = "pid"
grid1.TextMatrix(0, 1) = "ptype"
grid1.TextMatrix(0, 2) = "pcompany"

grid1.TextMatrix(0, 3) = "pmodelname"
grid1.TextMatrix(0, 4) = "pmrp"

Dim i As Integer
i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = rs!PID
grid1.TextMatrix(i, 1) = rs!ptype
grid1.TextMatrix(i, 2) = rs!pcompany

grid1.TextMatrix(i, 3) = rs!pmodelname
grid1.TextMatrix(i, 4) = rs!pmrp

i = i + 1
rs.MoveNext
Wend
End Sub
Private Sub cmdadd_Click()
If cmdadd.Caption = "add" Then
clear

'rs.MoveLast
'lblpid.Caption = Max(rs!PID) + 1

If rs.State = 1 Then rs.Close
rs.Open "select * from product", con, 3, 3
lblpid.Caption = 10
While Not rs.EOF
If Val(rs!PID) > Val(lblpid.Caption) Then lblpid.Caption = Val(rs!PID)
rs.MoveNext
Wend
lblpid.Caption = Val(lblpid.Caption) + 1

enableall
cmdadd.Caption = "save"
Else
rs.addnew
rs!PID = Val(lblpid.Caption)
rs!pcompany = combopcompany.Text
rs!ptype = comboptype.Text
rs!pmodelname = txtpmodelname.Text
rs!pmrp = txtpmrp.Text
rs.Update
cmdadd.Caption = "add"
rs.Close
rs.Open "select * from product", con, 3, 3
gridfill
clear
disableall
MsgBox "new product added", , "Done"
Unload Me
Me.Show
End If
End Sub
'Private Sub cmddel_Click()
'con.Execute "delete from product where pid=" & grid1.TextMatrix(grid1.RowSel, 0)
'rs.Close
'rs.Open "select * from product", con, 3, 3
'clear
'gridfill
'End Sub

Public Sub clear()
lblpid.Caption = ""
combopcompany.Text = ""
comboptype.Text = ""
txtpmodelname.Text = ""
txtpmrp.Text = ""
End Sub

Private Sub grid1_Click()
lblpid.Caption = grid1.TextMatrix(grid1.RowSel, 0)
comboptype.Text = grid1.TextMatrix(grid1.RowSel, 1)
combopcompany.Text = grid1.TextMatrix(grid1.RowSel, 2)

txtpmodelname.Text = grid1.TextMatrix(grid1.RowSel, 3)
txtpmrp.Text = grid1.TextMatrix(grid1.RowSel, 4)
End Sub

Public Sub enableall()
'txtpid.Enabled = True
combopcompany.Enabled = True
comboptype.Enabled = True
txtpmodelname.Enabled = True
txtpmrp.Enabled = True
End Sub

Public Sub disableall()
'txtpid.Enabled = False
combopcompany.Enabled = False
comboptype.Enabled = False
txtpmodelname.Enabled = False
txtpmrp.Enabled = False
End Sub

