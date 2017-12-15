VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "End"
      Height          =   615
      Left            =   9000
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Logout"
      Height          =   615
      Left            =   9000
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Generate Report"
      Height          =   735
      Left            =   5160
      TabIndex        =   5
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Assign Platform"
      Height          =   735
      Left            =   5160
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Queue Details"
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Platform Details"
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Train Details"
      Height          =   735
      Left            =   5160
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Choose Option"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form4.Show
Unload Me
End Sub

Private Sub Command3_Click()
Form5.Show
Unload Me
End Sub

Private Sub Command4_Click()
Form6.Show
Unload Me
End Sub

Private Sub Command5_Click()
Form7.Show
Unload Me
End Sub

Private Sub Command6_Click()
Form1.Show
Unload Me
End Sub

Private Sub Command7_Click()
End
End Sub
