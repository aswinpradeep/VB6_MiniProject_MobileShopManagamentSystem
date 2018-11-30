VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmstock 
   Caption         =   "STOCK"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   9870
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   2535
      Left            =   840
      TabIndex        =   0
      Top             =   2760
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4471
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
   Begin VB.Shape Shape2 
      Height          =   3015
      Left            =   720
      Top             =   2520
      Width           =   7935
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   2520
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "(in ascending order as per quantity)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "STOCK STATUS"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "select * from stock order by stqty", con, 3, 3

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

