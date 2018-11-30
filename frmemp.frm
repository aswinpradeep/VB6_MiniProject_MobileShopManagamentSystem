VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmemp 
   Caption         =   "EMPLOYEE"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14820
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   14820
   Begin MSComCtl2.DTPicker dtpempdob 
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      Top             =   3360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      _Version        =   393216
      Format          =   73072641
      CurrentDate     =   42981
   End
   Begin VB.TextBox txtempname 
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
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
      Left            =   1080
      Picture         =   "frmemp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "EDIT"
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
      Picture         =   "frmemp.frx":0C42
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   975
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
      Left            =   4320
      Picture         =   "frmemp.frx":1884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "SEARCH"
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
      Left            =   8520
      Picture         =   "frmemp.frx":24C6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   975
   End
   Begin VB.ComboBox combosearch 
      Height          =   315
      Left            =   6960
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   1335
      Left            =   1200
      TabIndex        =   2
      Top             =   4800
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   3
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
   Begin VB.Shape Shape3 
      Height          =   1455
      Left            =   360
      Top             =   6480
      Width           =   9855
   End
   Begin VB.Shape Shape2 
      Height          =   2895
      Left            =   960
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   1560
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblempid 
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "DOB"
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
      Left            =   1320
      TabIndex        =   10
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NEW EMPLOYEE DETAILS"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NAME"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE-ID"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   1680
      Width           =   2895
   End
End
Attribute VB_Name = "frmemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
txtempname.Enabled = True
dtpempdob.Enabled = True



If cmdadd.Caption = "ADD" Then

clear


If rs.State = 1 Then rs.Close
rs.Open "select * from employee", con, 3, 3
lblempid.Caption = 750
While Not rs.EOF
If Val(rs!empid) > Val(lblempid.Caption) Then lblempid.Caption = Val(rs!empid)
rs.MoveNext
Wend
lblempid.Caption = Val(lblempid.Caption) + 1

cmdadd.Caption = "SAVE"
Else
rs.addnew
rs.Fields(0) = Val(lblempid.Caption)
rs.Fields(1) = txtempname.Text
rs.Fields(2) = dtpempdob.Value

rs.Update

cmdadd.Caption = "ADD"
gridfill
clear
txtempname.Enabled = False
dtpempdob.Enabled = False
combofill
MsgBox "employee added", , "Done"
Unload Me
Me.Show
End If
End Sub


Private Sub Command7_Click()
clear
End Sub

Private Sub cmddel_Click()
'''
con.Execute "delete from logintbl where empid=" & lblempid.Caption
con.Execute "delete from employee where empid=" & lblempid.Caption
gridfill
End Sub

Private Sub cmdsearch_Click()

If combosearch.Text = "" Then
gridfill
Else

If rs.State = 1 Then rs.Close
rs.Open "select * from employee where empname='" & combosearch.Text & "'", con, 3, 3
Dim i As Integer
grid1.Rows = 1
grid1.TextMatrix(0, 0) = "empid"
grid1.TextMatrix(0, 1) = "empname"
grid1.TextMatrix(0, 2) = "empdob"

i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = rs.Fields(0)
grid1.TextMatrix(i, 1) = rs.Fields(1)
grid1.TextMatrix(i, 2) = rs.Fields(2)
i = i + 1
rs.MoveNext
Wend
End If
End Sub

Private Sub cmdedit_Click()
If lblempid.Caption = "" Then
MsgBox "please select something to edit", , "Warning"
Else


If cmdedit.Caption = "EDIT" Then
txtempname.Enabled = True
dtpempdob.Enabled = True
cmdedit.Caption = "SAVE"
Else
If rs.State = 1 Then rs.Close
rs.Open "select * from employee where empid=" & lblempid.Caption, con, 3, 3
txtempname.Enabled = True
dtpempdob.Enabled = True
rs.Fields(0) = Val(lblempid.Caption)
rs.Fields(1) = txtempname.Text
rs.Fields(2) = dtpempdob.Value
rs.Update

cmdedit.Caption = "EDIT"
gridfill
clear
txtempname.Enabled = False
dtpempdob.Enabled = False
combofill
MsgBox "editing done", , "Done"
End If
End If
End Sub

Private Sub Form_Load()

txtempname.Enabled = False
dtpempdob.Enabled = False

combofill

gridfill
End Sub

Private Sub grid1_Click()
lblempid.Caption = grid1.TextMatrix(grid1.RowSel, 0)
txtempname.Text = grid1.TextMatrix(grid1.RowSel, 1)
dtpempdob.Value = grid1.TextMatrix(grid1.RowSel, 2)
End Sub

Public Sub gridfill()
Dim i As Integer
If rs.State = 1 Then rs.Close
rs.Open "select * from employee", con, 3, 3
grid1.Rows = 1
grid1.TextMatrix(0, 0) = "empid"
grid1.TextMatrix(0, 1) = "empname"
grid1.TextMatrix(0, 2) = "empdob"

i = 1
While Not rs.EOF
grid1.Rows = grid1.Rows + 1
grid1.TextMatrix(i, 0) = rs.Fields(0)
grid1.TextMatrix(i, 1) = rs.Fields(1)
grid1.TextMatrix(i, 2) = rs.Fields(2)
i = i + 1
rs.MoveNext
Wend
End Sub

Public Sub clear()
lblempid.Caption = ""
txtempname.Text = ""
dtpempdob.Value = Date
End Sub



Public Sub combofill()
combosearch.clear
If rs.State = 1 Then rs.Close
rs.Open "select empname from employee", con, 3, 3
While Not rs.EOF
combosearch.AddItem (rs!empname)
rs.MoveNext
Wend
End Sub
