VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   1695
      Left            =   1440
      OleObjectBlob   =   "Form3.frx":0014
      TabIndex        =   20
      Top             =   7440
      Width           =   11415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\New folder\abc11.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "train"
      Top             =   9360
      Width           =   1260
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Back"
      Height          =   495
      Left            =   11520
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   495
      Left            =   10560
      TabIndex        =   18
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      Height          =   495
      Left            =   9000
      TabIndex        =   17
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   495
      Left            =   7440
      TabIndex        =   16
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   495
      Left            =   10080
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      DataField       =   "rest_time"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4740
      TabIndex        =   14
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      DataField       =   "destination"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   9840
      TabIndex        =   13
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      DataField       =   "Source"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4740
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      DataField       =   "dept_time"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   9840
      TabIndex        =   11
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataField       =   "arr_time"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4740
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "train_name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   9840
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "train_no"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4740
      TabIndex        =   8
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Rest Time"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Destination"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Source"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Depatrure Time"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Arrival Time"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Train Name"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Train No"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Train Details"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Private Sub Command2_Click()
Data1.Recordset.AddNew
Data1.Recordset!train_no = Val(Text1.Text)
Data1.Recordset!train_name = Text2.Text
Data1.Recordset!arr_time = Val(Text3.Text)
Data1.Recordset!dept_time = Val(Text4.Text)
Data1.Recordset!Source = Text5.Text
Data1.Recordset!destination = Text6.Text
Data1.Recordset!rest_time = Val(Text7.Text)
Data1.Recordset.Update
Data1.Refresh
End Sub

Private Sub Command3_Click()
Dim f As Integer
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF = False
f = Data1.Recordset!train_no
If f = Val(Text1.Text) Then
Data1.Recordset.Edit
Data1.Recordset!train_no = Val(Text1.Text)
Data1.Recordset!train_name = Text2.Text
Data1.Recordset!arr_time = Text3.Text
Data1.Recordset!dept_time = Text4.Text
Data1.Recordset!Source = Text5.Text
Data1.Recordset!destination = Text6.Text
Data1.Recordset!rest_time = Text7.Text
Data1.Recordset.Update
Data1.Refresh
Exit Do
End If
Data1.Recordset.MoveNext
Loop
End Sub

Private Sub Command4_Click()
Dim a As Integer
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF = False
a = Data1.Recordset!train_no
If a = Val(Text1.Text) Then
Data1.Recordset.Delete
Data1.Refresh
End If
Data1.Recordset.MoveNext
Loop
End Sub

Private Sub Command5_Click()
Form2.Show
Unload Me
End Sub

