VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpurchasereturn 
   Caption         =   "PURCHASE RETURN"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   14790
   Begin VB.CommandButton cmddel 
      Caption         =   "-LIST"
      Height          =   735
      Left            =   13200
      Picture         =   "frmpurchasereturn.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "+LIST"
      Height          =   615
      Left            =   13080
      Picture         =   "frmpurchasereturn.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpretdate 
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   3060
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16384001
      CurrentDate     =   42981
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RETURN"
      Height          =   735
      Left            =   3600
      Picture         =   "frmpurchasereturn.frx":2714
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   1215
      Left            =   7080
      TabIndex        =   7
      Top             =   4440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   5
      AllowUserResizing=   1
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
   Begin VB.TextBox txtretqty 
      Height          =   615
      Left            =   13680
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   1455
      Left            =   7080
      TabIndex        =   5
      Top             =   1920
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   5
      AllowUserResizing=   1
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
   Begin VB.ComboBox combopuid 
      Height          =   315
      Left            =   3840
      TabIndex        =   4
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      Height          =   4935
      Left            =   6600
      Top             =   1440
      Width           =   8535
   End
   Begin VB.Shape Shape3 
      Height          =   1695
      Left            =   6840
      Top             =   4320
      Width           =   8055
   End
   Begin VB.Shape Shape2 
      Height          =   2295
      Left            =   6840
      Top             =   1680
      Width           =   7935
   End
   Begin VB.Shape Shape1 
      Height          =   3615
      Left            =   480
      Top             =   1920
      Width           =   5655
   End
   Begin VB.Label lblsid 
      Alignment       =   2  'Center
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
      Left            =   3840
      TabIndex        =   15
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "supplier id"
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
      Left            =   720
      TabIndex        =   14
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblretid 
      Alignment       =   2  'Center
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
      Left            =   3840
      TabIndex        =   13
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "QTY"
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
      Left            =   12720
      TabIndex        =   10
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "purchase id"
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
      Left            =   720
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "return date"
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
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "return-id"
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
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PURCHASE RETURN"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmpurchasereturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub combopuid_Click()
If rs.State = 1 Then rs.Close
rs.Open "select sid from purchase where puid=" & combopuid.Text, con, 3, 3
lblsid.Caption = rs!sid
gridfill
End Sub

Private Sub Command1_Click()
If rs.State = 1 Then rs.Close
rs.Open "return1", con, 3, 3
rs.addnew
rs!retID = Val(lblretid.Caption)
rs!retdate = dtpretdate.Value
rs!puid = combopuid.Text
rs!sid = Val(lblsid.Caption)
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from returntemp where retid=" & Val(lblretid.Caption), con, 3, 3
While Not rs1.EOF
rs!retgtotal = rs!retgtotal + rs1!rettotal
rs1.MoveNext
Wend
rs.Update
MsgBox "return completed", , "Done"
Unload Me
Me.Show
End Sub

Private Sub cmdadd_Click()

If rs.State = 1 Then rs.Close
rs.Open "select stqty from stock where pid=" & grid1.TextMatrix(grid1.RowSel, 2)
If Val(txtretqty.Text) > rs!stqty Then
MsgBox "entered qty greater than available stock", , "warning"
txtretqty.Text = ""
Else

returntempfill
grid2fill
con.Execute "update stock set stqty=stqty-" & grid1.TextMatrix(grid1.RowSel, 3) & "where pid=" & grid1.TextMatrix(grid1.RowSel, 2)
gridfill
txtretqty.Text = ""
End If
End Sub

Private Sub cmddel_Click()
con.Execute "update stock set stqty=stqty+" & grid2.TextMatrix(grid2.RowSel, 3) & "where pid=" & grid2.TextMatrix(grid2.RowSel, 1)
con.Execute "delete from returntemp where pid=" & grid2.TextMatrix(grid2.RowSel, 1)
grid2fill
gridfill
End Sub

Private Sub Form_Load()

If rs.State = 1 Then rs.Close
rs.Open "select * from returntemp", con, 3, 3
lblretid.Caption = 500
While Not rs.EOF
If Val(rs!retID) > Val(lblretid.Caption) Then lblretid.Caption = Val(rs!retID)
rs.MoveNext
Wend
lblretid.Caption = Val(lblretid.Caption) + 1

If rs.State = 1 Then rs.Close
rs.Open "select distinct puid from purchasetemp", con, 3, 3
While Not rs.EOF
combopuid.AddItem (rs!puid)
rs.MoveNext
Wend
End Sub


Public Sub gridfill()
If rs.State = 1 Then rs.Close
rs.Open "select * from purchasetemp where puid=" & Val(combopuid.Text), con, 3, 3
grid1.Rows = 1
grid1.TextMatrix(0, 0) = "slno"
grid1.TextMatrix(0, 1) = "sid"
grid1.TextMatrix(0, 2) = "pid"
grid1.TextMatrix(0, 3) = "puqty"
grid1.TextMatrix(0, 4) = "puprice"
Dim i As Integer
i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = i
grid1.TextMatrix(i, 1) = rs!sid
grid1.TextMatrix(i, 2) = rs!PID
grid1.TextMatrix(i, 3) = rs!puqty
grid1.TextMatrix(i, 4) = rs!puprice
i = i + 1
rs.MoveNext
Wend
End Sub

Public Sub grid1_DblClick()
returntempfill
grid2fill
con.Execute "update stock set stqty=stqty-" & Val(txtretqty.Text) & "where pid=" & grid1.TextMatrix(grid1.RowSel, 2)
gridfill
End Sub


Public Sub grid2fill()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from returntemp where retid=" & Val(lblretid.Caption), con, 3, 3
grid2.Rows = 1
grid2.TextMatrix(0, 0) = "slno"
grid2.TextMatrix(0, 1) = "pid"
grid2.TextMatrix(0, 2) = "puprice"
grid2.TextMatrix(0, 3) = "retqty"
grid2.TextMatrix(0, 4) = "rettotal"
Dim j As Integer
j = 1
While Not rs1.EOF
grid2.Rows = grid1.Rows + 1
grid2.TextMatrix(j, 0) = j
grid2.TextMatrix(j, 1) = rs1!PID
grid2.TextMatrix(j, 2) = rs1!puprice
grid2.TextMatrix(j, 3) = rs1!retqty
grid2.TextMatrix(j, 4) = rs1!rettotal
j = j + 1
rs1.MoveNext
Wend
End Sub

Public Sub returntempfill()

If rs1.State = 1 Then rs1.Close
rs1.Open "returntemp", con, 3, 3
rs1.addnew
rs1!PID = grid1.TextMatrix(grid1.RowSel, 2)
rs1!puprice = grid1.TextMatrix(grid1.RowSel, 4)
rs1!retqty = Val(txtretqty.Text)
rs1!rettotal = (Val(rs1!puprice) * Val(txtretqty.Text))
rs1!retID = Val(lblretid.Caption)
rs1.Update
End Sub


Public Sub grid2_DblClick()
con.Execute "update stock set stqty=stqty+" & grid2.TextMatrix(grid2.RowSel, 3) & "where pid=" & grid2.TextMatrix(grid2.RowSel, 1)
con.Execute "delete from returntemp where pid=" & grid2.TextMatrix(grid2.RowSel, 1)
grid2fill
gridfill
End Sub





