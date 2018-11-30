VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmload 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4485
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmload.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmload.frx":000C
   ScaleHeight     =   4485
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   1800
      Top             =   2280
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2400
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmload.frx":E9D2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   4
      Top             =   3480
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AMIGOS MOBILES"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "loading . . . . "
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
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblper 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   5280
      TabIndex        =   1
      Top             =   2400
      Width           =   735
   End
End
Attribute VB_Name = "frmload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
lblper.Caption = ProgressBar1.Value
If ProgressBar1.Value = 100 Then
       Unload Me
       frmlogin.Show
       'frmmdi.Show
       End If
End Sub
