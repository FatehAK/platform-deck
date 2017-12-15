VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form4.frx":0000
      Height          =   2535
      Left            =   4320
      OleObjectBlob   =   "Form4.frx":0014
      TabIndex        =   14
      Top             =   5760
      Width           =   6375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\New folder\abc2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "platform"
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Back"
      Height          =   615
      Left            =   1440
      TabIndex        =   13
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   615
      Left            =   1440
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   615
      Left            =   9360
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   615
      Left            =   9360
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   615
      Left            =   9360
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      DataField       =   "train_no"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "Status"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "length"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "plat_no"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Train Number"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Status"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Platform length"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Platform No"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Platform Details"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset!plat_no = Val(Text1.Text)
Data1.Recordset!length = Val(Text2.Text)
Data1.Recordset!Status = Val(Text3.Text)
Data1.Recordset!train_no = Val(Text4.Text)
Data1.Recordset.Update
Data1.Refresh
End Sub

Private Sub Command2_Click()
Dim e As Integer
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF = False
e = Data1.Recordset!plat_no
If e = Val(Text1.Text) Then
Data1.Recordset.Edit
Data1.Recordset!length = Val(Text2.Text)
Data1.Recordset!Status = Val(Text3.Text)
Data1.Recordset!train_no = Val(Text4.Text)
Data1.Recordset.Update
Data1.Refresh
Exit Do
End If
Data1.Recordset.MoveNext
Loop
End Sub

Private Sub Command3_Click()
Dim a As Integer
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF = False
a = Data1.Recordset!plat_no
If a = Val(Text1.Text) Then
Data1.Recordset.Delete
Data1.Refresh
End If
Data1.Recordset.MoveNext
Loop
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command5_Click()
Form2.Show
Unload Me
End Sub
