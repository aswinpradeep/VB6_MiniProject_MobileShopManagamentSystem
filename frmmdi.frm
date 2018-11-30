VERSION 5.00
Begin VB.MDIForm frmmdi 
   BackColor       =   &H8000000C&
   Caption         =   "AMIGOS MOBILES"
   ClientHeight    =   8265
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15450
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   840
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   15390
      TabIndex        =   0
      Top             =   0
      Width           =   15450
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000009&
         Height          =   615
         Left            =   12240
         Picture         =   "frmmdi.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000009&
         Height          =   615
         Left            =   9840
         Picture         =   "frmmdi.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000009&
         Height          =   615
         Left            =   7560
         Picture         =   "frmmdi.frx":1C6C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   4680
         Picture         =   "frmmdi.frx":3FDE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   2280
         Picture         =   "frmmdi.frx":5B20
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "LOGOUT"
         Height          =   255
         Left            =   11520
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "FINANCE ANALYZER"
         Height          =   615
         Left            =   8880
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK STATUS"
         Height          =   495
         Left            =   6720
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SELL PRODUCTS"
         Height          =   495
         Left            =   3600
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PURCHASE PRODUCTS"
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   19935
      End
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu changepassword 
         Caption         =   "change password"
      End
      Begin VB.Menu logout 
         Caption         =   "Logout"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu transactions 
      Caption         =   "Transactions"
      Begin VB.Menu purchaseproducts 
         Caption         =   "Purchase products"
      End
      Begin VB.Menu sellproducts 
         Caption         =   "sell products"
      End
      Begin VB.Menu returnpurchasedproducts 
         Caption         =   "Return purchased products"
      End
      Begin VB.Menu replacesoldproducts 
         Caption         =   "Replace sold products"
      End
   End
   Begin VB.Menu maintain 
      Caption         =   "Maintain"
      Begin VB.Menu product 
         Caption         =   "Product"
      End
      Begin VB.Menu supplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu employee 
         Caption         =   "Employee"
      End
   End
   Begin VB.Menu stockanalyser 
      Caption         =   "stock analyser"
   End
   Begin VB.Menu financeanalyser 
      Caption         =   "Finance analyser"
   End
   Begin VB.Menu Reports 
      Caption         =   "Summary reports"
      Begin VB.Menu employeereport 
         Caption         =   "Employee report"
      End
      Begin VB.Menu productreport 
         Caption         =   "Product report"
      End
      Begin VB.Menu stockreport 
         Caption         =   "stock report"
      End
      Begin VB.Menu purchase 
         Caption         =   "purchase"
         Begin VB.Menu purchasereport 
            Caption         =   "Purchase report"
         End
         Begin VB.Menu purchasereturnreport 
            Caption         =   "purchase return report"
         End
      End
      Begin VB.Menu sales 
         Caption         =   "Sales"
         Begin VB.Menu salesreport 
            Caption         =   "Sales report"
         End
         Begin VB.Menu salesreplacementreport 
            Caption         =   "Sales replacement report"
         End
      End
   End
   Begin VB.Menu settings 
      Caption         =   "settings"
      Begin VB.Menu adduser 
         Caption         =   "adduser"
      End
   End
End
Attribute VB_Name = "frmmdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adduser_Click()
frmusr.Show
End Sub

Private Sub changepassword_Click()
frmretrive.Show
End Sub

Private Sub Command1_Click()
frmpurchase.Show
End Sub

Private Sub Command2_Click()
frmbill.Show
End Sub

Private Sub Command3_Click()
frmstock.Show
End Sub

Private Sub Command4_Click()
frmstatus.Show
End Sub

Private Sub Command5_Click()
Unload Me
frmlogin.Show
End Sub

Private Sub employee_Click()
frmemp.Show
End Sub

Private Sub employeereport_Click()
eemployeereport.Show
End Sub

Private Sub exit_Click()
x = MsgBox("Are you sure?", vbYesNo, "Warning")
If x = vbYes Then
Unload Me
Else
'Me.Show
'Me.Hide
Unload Me
Me.Show
End If
End Sub

Private Sub financeanalyser_Click()
frmstatus.Show
End Sub

Private Sub logout_Click()
Unload Me
frmlogin.Show
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub product_Click()
frmproduct.Show
End Sub

Private Sub productreport_Click()
pproductreport.Show
End Sub

Private Sub purchaseproducts_Click()
frmpurchase.Show
End Sub

Private Sub purchasereport_Click()
ppurchasereport.Show
End Sub

Private Sub purchasereturnreport_Click()
ppurchasereturnreport.Show
End Sub

Private Sub replacesoldproducts_Click()
frmreplacement.Show
End Sub

Private Sub returnpurchasedproducts_Click()
frmpurchasereturn.Show
End Sub

Private Sub salesreplacementreport_Click()
rreplacementreport.Show
End Sub

Private Sub salesreport_Click()
ssalesreport.Show
End Sub

Private Sub sellproducts_Click()
frmbill.Show
End Sub

Private Sub stockanalyser_Click()
frmstock.Show
End Sub

Private Sub stockreport_Click()
sstockreport.Show
End Sub

Private Sub supplier_Click()
frmsupplier.Show
End Sub
