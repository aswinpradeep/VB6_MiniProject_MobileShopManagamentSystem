VERSION 5.00
Begin VB.Form frmusr 
   Caption         =   "USER"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   8550
   Begin VB.TextBox txtusr 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   3045
      Width           =   2895
   End
   Begin VB.TextBox txtpass1 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   3975
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD USER"
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
      Left            =   3360
      Picture         =   "frmusr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1575
   End
   Begin VB.ComboBox Comboid 
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtpass2 
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      Height          =   4215
      Left            =   600
      Top             =   1560
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   2640
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ADD NEW USER"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "USER-NAME"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   3045
      Width           =   2895
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   4035
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE I-D"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "REPEAT PASSWORD"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   5040
      Width           =   2895
   End
End
Attribute VB_Name = "frmusr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If rs.State = 1 Then rs.Close
rs.Open "logintbl", con, 3, 3

While Not rs.EOF
        Dim v1 As Boolean
        v1 = True
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select distinct empid from logintbl", con, 3, 3
        While Not rs1.EOF
            If Comboid.Text = rs1!empid Then
                MsgBox "employee already exist", , "warning"
                'rs1.MoveLast
                rs.MoveLast
                v1 = False
            End If
            rs1.MoveNext
            Wend
        If v1 = True Then
            If rs1.State = 1 Then rs1.Close
        rs1.Open "select distinct loginuserid from logintbl", con, 3, 3
        While Not rs1.EOF
            If txtusr.Text = rs1!loginuserid Then
                MsgBox "user already exist", , "warning"
                'rs1.MoveLast
                rs.MoveLast
                v1 = False
            End If
            rs1.MoveNext
            Wend
                 
        End If
        If v1 = True Then
        
        If txtpass1.Text <> txtpass2.Text Then
            MsgBox "passwords dosent match", , "warning"
            rs.MoveLast
            v1 = False
        End If
        
        End If
    If v1 = True Then
      
      If rs.State = 1 Then rs.Close
         rs.Open "logintbl", con, 3, 3
         rs.addnew
        rs.Fields(0) = txtusr.Text
        rs.Fields(1) = txtpass1.Text
        rs.Fields(2) = Comboid.Text
        rs.Update
        MsgBox "new user added successfully", , "Done"
        Unload Me
        frmlogin.Show
        End If
rs.MoveNext
Wend
End Sub
Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "select empid from employee", con, 3, 3
While Not rs.EOF
Comboid.AddItem (rs!empid)
rs.MoveNext
Wend
End Sub


Public Sub clear()
txtusr.Text = ""
txtpass1.Text = ""
txtpass2.Text = ""
Comboid.Text = ""
End Sub

