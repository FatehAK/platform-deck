VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   615
      Left            =   6240
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   7560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Admin Login"
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "admin" And Text2.Text = "password" Then
Form2.Show
Unload Me
Else
MsgBox "Incorrect username or password"
Form1.Show
Unload Me
End If
End Sub

