VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmretrive 
   Caption         =   "PASSWORD UTILITY"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form7"
   ScaleHeight     =   8850
   ScaleWidth      =   10425
   Begin MSComCtl2.DTPicker dtpicker1 
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   3000
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   42981
   End
   Begin VB.ComboBox comempid 
      Height          =   315
      Left            =   5280
      TabIndex        =   4
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
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
      Left            =   4080
      Picture         =   "frmretrive.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox pass1 
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox pass2 
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   6840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3840
      Picture         =   "frmretrive.frx":0C42
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      Height          =   3135
      Left            =   720
      Top             =   4800
      Width           =   8175
   End
   Begin VB.Shape Shape2 
      Height          =   2415
      Left            =   720
      Top             =   1200
      Width           =   8175
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   2520
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RETRIVE USERNAME AND PASSWORD"
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
      TabIndex        =   11
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label2 
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
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "DATE OF BIRTH"
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
      Left            =   960
      TabIndex        =   9
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "USER-ID IS:"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label5 
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
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "ENTER NEW PASSWORD"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   6000
      Width           =   3135
   End
   Begin VB.Label Label7 
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
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   6960
      Width           =   3135
   End
End
Attribute VB_Name = "frmretrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ls = 0
If rs.State = 1 Then rs.Close
rs.Open "select * from employee", con, 3, 3
        While Not rs.EOF
            If comempid.Text = rs!empid Then
                If dtpicker1.Value = rs!empdob Then
                    pass1.Visible = True
                    pass2.Visible = True
                    ls = 1
                End If
              
            End If
            rs.MoveNext
        Wend
If ls = 0 Then
MsgBox "invalid credentials", , "warning"
End If
End Sub

Private Sub Command2_Click()
If pass1.Text = pass2.Text Then
If rs.State = 1 Then rs.Close
rs.Open "select * from logintbl where empid='" & comempid.Text & "'", con, 3, 3
rs!loginpass = pass1.Text
rs.Update
MsgBox "updated successfully", , "Done"
Unload Me
frmlogin.Show
Else
MsgBox "passwords dosent match", , "Warning"
End If
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "select distinct empid from employee", con, 3, 3
        While Not rs.EOF
            comempid.AddItem (rs!empid)
            rs.MoveNext
        Wend
End Sub

