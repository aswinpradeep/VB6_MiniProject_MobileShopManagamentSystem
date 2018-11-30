VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreplacement 
   Caption         =   "REPLACEMENT"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   15765
   Begin VB.TextBox txtpqty 
      Height          =   285
      Left            =   10800
      TabIndex        =   20
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "check"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      Picture         =   "frmsalesreturn.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   1335
   End
   Begin VB.ComboBox comboreason 
      Height          =   315
      ItemData        =   "frmsalesreturn.frx":1B42
      Left            =   10800
      List            =   "frmsalesreturn.frx":1B4C
      TabIndex        =   15
      Top             =   5160
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtpbilldate 
      Height          =   615
      Left            =   4200
      TabIndex        =   13
      Top             =   4800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   43069
   End
   Begin MSComCtl2.DTPicker dtprepdate 
      Height          =   615
      Left            =   4200
      TabIndex        =   11
      Top             =   2355
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   43069
   End
   Begin VB.ComboBox combopid 
      Height          =   315
      Left            =   10800
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox combobillid 
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Top             =   3270
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "replace&printbill"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7920
      Picture         =   "frmsalesreturn.frx":1B6A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lbls 
      Height          =   255
      Left            =   12720
      TabIndex        =   24
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "stock available"
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
      Left            =   10680
      TabIndex        =   23
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   4215
      Left            =   6960
      Top             =   1560
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      Height          =   5655
      Left            =   480
      Top             =   1440
      Width           =   6255
   End
   Begin VB.Label lblmax 
      Height          =   255
      Left            =   12840
      TabIndex        =   22
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "max"
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
      Left            =   12000
      TabIndex        =   21
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "pqty"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   19
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label lblbillname 
      Alignment       =   2  'Center
      Caption         =   "1252"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   18
      Top             =   3885
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "billname"
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
      Left            =   720
      TabIndex        =   17
      Top             =   4290
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "reason"
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
      Left            =   7440
      TabIndex        =   14
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "billdate"
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
      Left            =   720
      TabIndex        =   12
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "return date"
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
      Left            =   720
      TabIndex        =   10
      Top             =   2550
      Width           =   2655
   End
   Begin VB.Label lblpname 
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
      Height          =   375
      Left            =   10800
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "product-name"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   5
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "product-id"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   4
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3420
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "SALES REPLACEMENT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label lblrepid 
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
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "replacement-id"
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
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
End
Attribute VB_Name = "frmreplacement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub combobillid_click()
If rs.State = 1 Then rs.Close
rs.Open "select billdate,billname from bill where billid=" & combobillid.Text, con, 3, 3
dtpbilldate.Value = rs!billdate
lblbillname.Caption = rs!billname


If rs.State = 1 Then rs.Close
rs.Open "select pid from billtemp where billid=" & combobillid.Text, con, 3, 3
While Not rs.EOF
    combopid.AddItem (rs!PID)
rs.MoveNext
Wend
End Sub



Private Sub combopid_Click()
If rs.State = 1 Then rs.Close
rs.Open "select pmodelname from product where pid=" & combopid.Text, con, 3, 3

lblpname.Caption = rs!pmodelname

If rs.State = 1 Then rs.Close
rs.Open "select qty from billtemp where pid=" & combopid.Text & "and billid=" & combobillid.Text, con, 3, 3
lblmax.Caption = rs!qty

If rs.State = 1 Then rs.Close
rs.Open "select stqty from stock where pid=" & combopid.Text, con, 3, 3
lbls.Caption = rs!stqty
'Dim x As Integer
'Dim i As Integer
'i = 1
'x = rs!qty
'While i < x
'combobil
End Sub

Private Sub Command1_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from replacement", con, 3, 3
rs.addnew
rs!repid = Val(lblrepid.Caption)
rs!repdate = dtprepdate.Value
rs!billid = combobillid.Text
rs!billname = lblbillname.Caption
rs!PID = combopid.Text
rs!pname = lblpname.Caption
rs!pqty = txtpqty.Text
rs!reason = comboreason.Text
rs.Update
con.Execute "update stock set stqty=stqty-1 where pid=" & combopid.Text

DataEnvironment2.Connection1.Open
DataEnvironment2.Command1 lblbillname.Caption
DataReport2.Show

MsgBox "replacement done and stock updated", , "Done"
Unload Me
Me.Show


End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
If dtpbilldate.Value + 30 > Date Then
MsgBox "eligible for replacement", , "Proceed"
combopid.Enabled = True
comboreason.Enabled = True
Else
MsgBox "notelegible", , "warning"
combobillid.Text = ""

End If
End Sub

Private Sub Command4_Click()



End Sub

Private Sub dtpbilldate_Click()
If dtpbilldate.Value > dtprepdate.Value + 15 Then
MsgBox "not eligible for replacement", , "warning"
End If
End Sub

Private Sub Form_Load()
combopid.Enabled = False
comboreason.Enabled = False

'dtpbilldate.Value = Date + 30
If rs.State = 1 Then rs.Close
rs.Open "select * from replacement", con, 3, 3
lblrepid.Caption = 1250
While Not rs.EOF
If Val(rs!repid) > Val(lblrepid.Caption) Then lblrepid.Caption = Val(rs!repid)
rs.MoveNext
Wend
lblrepid.Caption = Val(lblrepid.Caption) + 1

If rs.State = 1 Then rs.Close
rs.Open "select billid from bill", con, 3, 3
While Not rs.EOF
    combobillid.AddItem (rs!billid)
rs.MoveNext
Wend


    
End Sub

Private Sub txtpqty_Change()
If Val(txtpqty.Text) > Val(lblmax.Caption) Or Val(txtpqty.Text) > Val(lbls.Caption) Then
MsgBox "invalid entry,recheck qty entered", , "Warning"
txtpqty.Text = ""
End If
End Sub
