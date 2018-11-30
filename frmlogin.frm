VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "LOGIN"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form6"
   ScaleHeight     =   6375
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtusr 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   3
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton Cmdlogin 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      Picture         =   "frmlogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      Picture         =   "frmlogin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FORGET ID/PASSWORD?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      Picture         =   "frmlogin.frx":1084
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   2520
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "USER-ID"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "PASSWORD"
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
      TabIndex        =   6
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "LOGIN FORM"
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
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loginsucceed As Boolean
Private Sub Cmdlogin_Click()
If rs.State = 1 Then rs.Close
rs.Open "logintbl", con, 3, 3
        While Not rs.EOF
            If txtusr = Trim(rs.Fields(0)) Or 1 = 1 Then
                 If txtpass = Trim(rs.Fields(1)) Or 1 = 1 Then
                      LoginSucceeded = True
                      rs.MoveLast
                 End If
            End If
        rs.MoveNext
        Wend
If LoginSucceeded = True Then
Unload Me
frmmdi.Show
Else
MsgBox "error!!!!invalid login credentials", , "Warning"
End If
End Sub

Private Sub Command1_Click()
Unload Me
frmretrive.Show
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
LoginSucceeded = False
End Sub

