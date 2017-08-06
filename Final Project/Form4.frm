VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11955
   LinkTopic       =   "Form4"
   ScaleHeight     =   6600
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   975
      Left            =   5520
      TabIndex        =   16
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   2400
      TabIndex        =   15
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   14
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   2400
      TabIndex        =   13
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   12
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   11
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   9
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   8
      Top             =   600
      Width           =   1935
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
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
rst.Close
rst.Open

rst.AddNew
rst.Fields(0) = Text1(0).Text
rst.Fields(1) = Text1(8).Text
rst.Fields(2) = Text1(7).Text
rst.Fields(3) = Text1(6).Text
rst.Fields(4) = Text1(5).Text
rst.Fields(5) = Text1(4).Text
rst.Fields(6) = Text1(3).Text
rst.Fields(7) = Text1(2).Text
rst.Update
rst.Close
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub



Private Sub Form_Load()

conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "Db1.mdb;Mode=read|write"
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
End Sub
