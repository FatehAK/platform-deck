VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\New folder\abc2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "platform"
      Top             =   5880
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   495
      Left            =   9360
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   9360
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Submit"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Assigned Platform"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Train Number"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Assign Platfrom"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Show
Unload Me
End Sub

Private Sub Command2_Click()
Dim b, c As Integer
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF = False
If Data1.Recordset!Status = 0 Or Data1.Recordset!Status = " " Then
b = Data1.Recordset!plat_no
c = Data1.Recordset!length
Text2.Text = b
Data1.Recordset.Delete
Data1.Recordset.AddNew
Data1.Recordset!plat_no = b
Data1.Recordset!length = c
Data1.Recordset!Status = 1
Data1.Recordset!train_no = Val(Text1.Text)
Data1.Recordset.Update
Data1.Refresh
Exit Sub
Exit Do
End If
Data1.Recordset.MoveNext
Loop
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command4_Click()
Form2.Show
Unload Me
End Sub
