VERSION 5.00
Begin VB.Form req 
   Caption         =   "Stock Requisition"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   LinkTopic       =   "Form4"
   ScaleHeight     =   7305
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   2400
      TabIndex        =   17
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Return"
      Height          =   612
      Left            =   2520
      TabIndex        =   16
      Top             =   6480
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   2400
      TabIndex        =   15
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   2400
      TabIndex        =   14
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   12
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Expected delivery date"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Total"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Rate"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Unit"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Stock id"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Supplier"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Re date"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Req no."
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "req"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()

On Error GoTo Err

With cmd
.ActiveConnection = conn
.CommandText = "Insert into Requisition values('" & Text1(0).Text & "','" & CDate(Text1(1).Text) & "','" & Text1(2).Text & "','" & Text1(3).Text & "','" & Text1(4).Text & "','" & Text2.Text & "','" & Text1(6).Text & "','" & CDate(Text1(7).Text) & "')"
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data

MsgBox ("data inserted successfully")
.CommandText = "select * from Stock"
rst.Close
rst.Open cmd

Err:
MsgBox ("We Regret The Inconvenience.")
Exit Sub

End With

End Sub



Private Sub Command2_Click()
conn.Close
Unload Me
stockcheck.Show
End Sub

Private Sub Form_Load()

conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "DB1.mdb;Mode=read|write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Requisition"
.CommandType = adCmdText
End With

With rst
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd

End With

Text1(3).Text = stockcheck.Text1.Text

End Sub

Private Sub Text2_LostFocus()
Text1(6).Text = Text1(4).Text * Text2.Text
End Sub
