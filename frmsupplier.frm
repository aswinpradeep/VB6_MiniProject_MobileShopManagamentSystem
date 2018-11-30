VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsupplier 
   Caption         =   "SUPPLIER"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
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
      Left            =   10320
      Picture         =   "frmsupplier.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtsmobno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   14
      Top             =   5070
      Width           =   2415
   End
   Begin VB.TextBox txtsname 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   8
      Top             =   2970
      Width           =   2415
   End
   Begin VB.TextBox txtsplace 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   7
      Top             =   4020
      Width           =   2415
   End
   Begin VB.TextBox txtsmailid 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   6
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "R"
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
      Left            =   14040
      TabIndex        =   5
      Top             =   5520
      Width           =   375
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
      Left            =   12240
      TabIndex        =   4
      Text            =   "Search"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "del"
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
      Left            =   8880
      Picture         =   "frmsupplier.frx":0C42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
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
      Left            =   7320
      Picture         =   "frmsupplier.frx":1884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   2055
      Left            =   7080
      TabIndex        =   0
      Top             =   2400
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   5
      AllowUserResizing=   1
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
   Begin VB.Shape Shape4 
      Height          =   735
      Left            =   5520
      Top             =   240
      Width           =   3855
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   12120
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      Height          =   1455
      Left            =   7080
      Top             =   5040
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      Height          =   5415
      Left            =   240
      Top             =   1680
      Width           =   6375
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
      Height          =   735
      Left            =   3840
      TabIndex        =   16
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "mobno"
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
      Left            =   480
      TabIndex        =   13
      Top             =   5190
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "sid"
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
      Left            =   480
      TabIndex        =   12
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "sname"
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
      Left            =   480
      TabIndex        =   11
      Top             =   3090
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "splace"
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
      Left            =   480
      TabIndex        =   10
      Top             =   4140
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "smailid"
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
      Left            =   480
      TabIndex        =   9
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "SUPPLIER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdadd_Click()
If cmdadd.Caption = "add" Then
clear
If rs.State = 1 Then rs.Close

rs.Open "select * from supplier", con, 3, 3
lblsid.Caption = 1
While Not rs.EOF
If Val(rs!sid) > Val(lblsid.Caption) Then
lblsid.Caption = Val(rs!sid)
End If
rs.MoveNext
Wend
lblsid.Caption = Val(lblsid.Caption) + 1
'rs.MoveLast
'lblsid.Caption = Val(rs!sid) + 1
enableall
cmdadd.Caption = "save"
Else
'If rs.State = 1 Then rs.Close
'rs.Open "select * from supplier", con, 3, 3
'rs.MoveLast
'lblsid.Caption = Val(rs!sid) + 1

'rs.MoveLast
'lblsid.Caption = Val(rs!sid) + 1
rs.addnew
rs!sid = Val(lblsid.Caption)
rs!sname = txtsname.Text
rs!splace = txtsplace.Text
rs!smobno = txtsmobno.Text
rs!smailid = txtsmailid.Text
rs.Update
cmdadd.Caption = "add"
rs.Close
rs.Open "select * from supplier", con, 3, 3
gridfill
clear
disableall
MsgBox "supplier added successfully", , "Done"
Unload Me
Me.Show
End If
End Sub

Private Sub cmddel_Click()
con.Execute "delete from supplier where sid=" & Val(lblsid.Caption)
rs.Close
rs.Open "select * from supplier", con, 3, 3
clear
gridfill
End Sub



Private Sub cmdedit_Click()
If cmdedit.Caption = "edit" Then
'clear
enableall
cmdedit.Caption = "save"
Else
If rs.State = 1 Then rs.Close
rs.Open "select * from supplier where sid=" & Val(lblsid.Caption), con, 3, 3
'rs.addnew
rs!sid = Val(lblsid.Caption)
rs!sname = txtsname.Text
rs!splace = txtsplace.Text
rs!smobno = txtsmailid.Text
rs!smailid = txtsmobno.Text
rs.Update
cmdedit.Caption = "edit"
rs.Close
rs.Open "select * from supplier", con, 3, 3
gridfill
clear
disableall
End If
End Sub



Private Sub Combo1_Change()
If rs.State = 1 Then rs.Close
rs.Open "select * from supplier", con, 3, 3
grid1.Rows = 1
grid1.TextMatrix(0, 0) = "sid"
grid1.TextMatrix(0, 1) = "sname"
grid1.TextMatrix(0, 2) = "splace"
grid1.TextMatrix(0, 3) = "smobno"
grid1.TextMatrix(0, 4) = "smailid"
Dim i As Integer
i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = rs!sid
grid1.TextMatrix(i, 1) = rs!sname
grid1.TextMatrix(i, 2) = rs!splace
grid1.TextMatrix(i, 3) = rs!smobno
grid1.TextMatrix(i, 4) = rs!smailid
i = i + 1
rs.MoveNext
Wend
End Sub
End Sub

Private Sub cmdreset_Click()
clear
disableall
cmdadd.Caption = "add"
cmdedit.Caption = "edit"
combosearch.Text = "Search"
gridfill
End Sub

Private Sub combosearch_Change()
If rs.State = 1 Then rs.Close
rs.Open "select * from supplier where sname like'" & combosearch.Text & "%'", con, 3, 3
grid1.Rows = 1
grid1.TextMatrix(0, 0) = "sid"
grid1.TextMatrix(0, 1) = "sname"
grid1.TextMatrix(0, 2) = "splace"
grid1.TextMatrix(0, 3) = "smobno"
grid1.TextMatrix(0, 4) = "smailid"
Dim i As Integer
i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = rs!sid
grid1.TextMatrix(i, 1) = rs!sname
grid1.TextMatrix(i, 2) = rs!splace
grid1.TextMatrix(i, 3) = rs!smobno
grid1.TextMatrix(i, 4) = rs!smailid
i = i + 1
rs.MoveNext
Wend
End Sub

Private Sub Form_Load()
'If rs.State = 1 Then rs.Close
'rs.Open "select * from supplier", con, 3, 3
'rs.MoveLast
'lblsid.Caption = Val(rs!sid) + 1


disableall
gridfill
End Sub

Public Sub gridfill()
If rs.State = 1 Then rs.Close
rs.Open "select * from supplier", con, 3, 3
grid1.Rows = 1
grid1.TextMatrix(0, 0) = "sid"
grid1.TextMatrix(0, 1) = "sname"
grid1.TextMatrix(0, 2) = "splace"
grid1.TextMatrix(0, 3) = "smobno"
grid1.TextMatrix(0, 4) = "smailid"
Dim i As Integer
i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = rs!sid
grid1.TextMatrix(i, 1) = rs!sname
grid1.TextMatrix(i, 2) = rs!splace
grid1.TextMatrix(i, 3) = rs!smobno
grid1.TextMatrix(i, 4) = rs!smailid
i = i + 1
rs.MoveNext
Wend
End Sub

Private Sub grid1_Click()
lblsid.Caption = grid1.TextMatrix(grid1.RowSel, 0)
txtsname.Text = grid1.TextMatrix(grid1.RowSel, 1)
txtsplace.Text = grid1.TextMatrix(grid1.RowSel, 2)
txtsmobno.Text = grid1.TextMatrix(grid1.RowSel, 3)
txtsmailid.Text = grid1.TextMatrix(grid1.RowSel, 4)

End Sub

Public Sub clear()
lblsid.Caption = ""
txtsname.Text = ""
txtsplace.Text = ""
txtsmobno.Text = ""
txtsmailid.Text = ""

End Sub

Public Sub disableall()
'txtsid.Enabled = False
txtsname.Enabled = False
txtsplace.Enabled = False
txtsmobno.Enabled = False
txtsmailid.Enabled = False
End Sub
Public Sub enableall()
'txtsid.Enabled = True
txtsname.Enabled = True
txtsplace.Enabled = True
txtsmobno.Enabled = True
txtsmailid.Enabled = True
End Sub
