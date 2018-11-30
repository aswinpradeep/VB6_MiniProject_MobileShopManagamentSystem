VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstatus 
   Caption         =   "FINANCE ANALYZER"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.TextBox txtemr 
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
      Left            =   12360
      TabIndex        =   16
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtems 
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
      Left            =   12360
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtp2 
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      Format          =   73007105
      CurrentDate     =   43068
   End
   Begin VB.CommandButton Command1 
      Caption         =   "calculate"
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
      Left            =   5640
      Picture         =   "frmstatus.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      Format          =   73007105
      CurrentDate     =   43068
   End
   Begin VB.Shape Shape3 
      Height          =   1215
      Left            =   2640
      Top             =   7320
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   2655
      Left            =   8160
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      Height          =   4455
      Left            =   240
      Top             =   2640
      Width           =   13935
   End
   Begin VB.Label Label8 
      Caption         =   "EXPENDITURE CALCULATOR"
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
      Left            =   3840
      TabIndex        =   17
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "extra money recieved"
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
      Left            =   8760
      TabIndex        =   14
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "extra money spend"
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
      Left            =   8640
      TabIndex        =   13
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "overall profit/loss"
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
      Left            =   3000
      TabIndex        =   12
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "tax amount"
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
      Left            =   600
      TabIndex        =   11
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "money recieved on sales"
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
      Left            =   600
      TabIndex        =   10
      Top             =   5145
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "money recieved on return"
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
      Left            =   600
      TabIndex        =   9
      Top             =   4050
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "money spend on purchase"
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
      Left            =   480
      TabIndex        =   8
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label lbltotal 
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
      Left            =   7800
      TabIndex        =   5
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Label lbltax 
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
      Left            =   4800
      TabIndex        =   4
      Top             =   6180
      Width           =   2895
   End
   Begin VB.Label lblsales 
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
      Left            =   4800
      TabIndex        =   3
      Top             =   5070
      Width           =   2895
   End
   Begin VB.Label lblreturn 
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
      Left            =   4800
      TabIndex        =   2
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label lblpurchase 
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
      Left            =   4800
      TabIndex        =   1
      Top             =   2835
      Width           =   2895
   End
End
Attribute VB_Name = "frmstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

txtems.Text = ""
txtemr.Text = ""

lblpurchase.Caption = 0
If rs.State = 1 Then rs.Close
'rs.Open "SELECT pugtotal from purchase where pudate='" & dtp.Value & "'", con, 3, 3
rs.Open "SELECT pugtotal from purchase where pudate between '" & dtp.Value & "' and '" & dtp2.Value & "'", con, 3, 3
While Not rs.EOF
lblpurchase.Caption = Val(lblpurchase.Caption) + Val(rs!pugtotal)
rs.MoveNext
Wend

lblreturn.Caption = 0
If rs.State = 1 Then rs.Close
'rs.Open "SELECT retgtotal from return1 where retdate='" & dtp.Value & "'", con, 3, 3
rs.Open "SELECT retgtotal from return1 where retdate between '" & dtp.Value & "' and '" & dtp2.Value & "'", con, 3, 3
While Not rs.EOF
lblreturn.Caption = Val(lblreturn.Caption) + Val(rs!retgtotal)
rs.MoveNext
Wend

lblsales.Caption = 0
If rs.State = 1 Then rs.Close
'rs.Open "SELECT billtotal from bill where billdate='" & dtp.Value & "'", con, 3, 3
rs.Open "SELECT billtotal from bill where billdate between '" & dtp.Value & "' and '" & dtp2.Value & "'", con, 3, 3
While Not rs.EOF
lblsales.Caption = Val(lblsales.Caption) + rs!billtotal
rs.MoveNext
Wend

lbltax.Caption = 0
If rs.State = 1 Then rs.Close
'rs.Open "SELECT billtax from bill where billdate='" & dtp.Value & " '", con, 3, 3
rs.Open "SELECT billtax from bill where billdate between '" & dtp.Value & "' and '" & dtp2.Value & "'", con, 3, 3
While Not rs.EOF
lbltax.Caption = Val(lbltax.Caption) + rs!billtax
rs.MoveNext
Wend

lbltotal.Caption = Val(lblsales.Caption) + Val(lblreturn.Caption) - Val(lbltax.Caption) - Val(lblpurchase.Caption) + Val(txtemr.Text) - Val(txtems.Text)
End Sub






'Private Sub lbltax_Change()
'lbltotal.Caption = Val(lblsales.Caption) + Val(lblreturn.Caption) - Val(lbltax.Caption) - Val(lblpurchase.Caption)
'End Sub

Private Sub Form_Load()
dtp.Value = Date
dtp2.Value = Date
End Sub

Private Sub txtemr_Change()
lbltotal.Caption = Val(lblsales.Caption) + Val(lblreturn.Caption) - Val(lbltax.Caption) - Val(lblpurchase.Caption) + Val(txtemr.Text) - Val(txtems.Text)
End Sub

Private Sub txtems_Change()
lbltotal.Caption = Val(lblsales.Caption) + Val(lblreturn.Caption) - Val(lbltax.Caption) - Val(lblpurchase.Caption) + Val(txtemr.Text) - Val(txtems.Text)
End Sub
