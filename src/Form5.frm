VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form5.frx":0000
      Height          =   2295
      Left            =   3120
      OleObjectBlob   =   "Form5.frx":0014
      TabIndex        =   16
      Top             =   6240
      Width           =   8295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\New folder\abc333.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "queue11"
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Back"
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   495
      Left            =   7680
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   7680
      TabIndex        =   12
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   7680
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "arr_time"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      DataField       =   "rep_time"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "date"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "train_no"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "que_no"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Arrival Time"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Reporting Time"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Date"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Train Number"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Queue Number"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Queue Details"
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset!que_no = Val(Text1.Text)
Data1.Recordset!train_no = Val(Text2.Text)
Data1.Recordset!Date = Text3.Text
Data1.Recordset!arr_time = Text4.Text
Data1.Recordset!rep_time = Text5.Text
Data1.Recordset.Update
Data1.Refresh
End Sub

Private Sub Command2_Click()
Dim a As Integer
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF = False
a = Data1.Recordset!que_no
If a = Val(Text1.Text) Then
Data1.Recordset.Edit
Data1.Recordset!que_no = Val(Text1.Text)
Data1.Recordset!train_no = Val(Text2.Text)
Data1.Recordset!Date = Text3.Text
Data1.Recordset!arr_time = Text4.Text
Data1.Recordset!rep_time = Text5.Text
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
a = Data1.Recordset!que_no
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
Text5.Text = ""
End Sub

Private Sub Command5_Click()
Form2.Show
Unload Me
End Sub

