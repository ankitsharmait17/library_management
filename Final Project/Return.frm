VERSION 5.00
Begin VB.Form Retur 
   Caption         =   "Return Document"
   ClientHeight    =   7536
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   9552
   LinkTopic       =   "Form1"
   ScaleHeight     =   8652
   ScaleWidth      =   16176
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Confirm"
      Height          =   492
      Left            =   5160
      TabIndex        =   15
      Top             =   960
      Width           =   2412
   End
   Begin VB.TextBox Text2 
      Height          =   492
      Left            =   2520
      TabIndex        =   14
      Top             =   960
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return The Document"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5160
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   2520
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   2520
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   2520
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Fine Due To Late Submission (If Any)"
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   480
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   480
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Submission Date"
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Expected Submission Date"
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Issue Date"
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Issue ID"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Member ID"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Retur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim rst1 As New ADODB.Recordset

Private Sub Command1_Click()
If Text1(6).Text > 0 Then
MsgBox ("A Fine of Rupees " & Text1(6).Text & " is charged for late submission. Charging 2 rupees per day.")
End If
Do While Not rst.EOF
If rst.Fields(0) = Text2.Text Then
Text1(5).Text = rst1.Fields(3)
Exit Do
End If
rst.MoveNext
Loop
rst.Delete
rst.Requery
rst.Update
MsgBox ("Returned!")
conn.Close
Unload Me
End Sub

Private Sub Command2_Click()
Dim d1 As Date
Dim d2 As Date
Dim i As Integer
Dim c As Integer
'On Error GoTo err

c = 0

Do While Not rst.EOF
If rst.Fields(0) = Text2.Text Then
c = 1
Label1(2).Visible = True
Label1(3).Visible = True
Label1(4).Visible = True
Label1(5).Visible = True
Label1(6).Visible = True
Text1(2).Visible = True
Text1(3).Visible = True
Text1(4).Visible = True
Text1(5).Visible = True
Text1(6).Visible = True
Command1.Enabled = True
Command1.Visible = True
Exit Do
End If
rst.MoveNext
Loop
If c = 1 Then
Command2.Visible = False
Text1(2).Text = rst.Fields(1)
Text1(3).Text = rst.Fields(4)
d1 = CDate(rst.Fields(4))
rst.Fields(5) = DateValue(Now)
Text1(4).Text = rst.Fields(5)
d2 = CDate(rst.Fields(5))
i = DateDiff("d", d1, d2, vbMonday, vbFirstJan1)
If i < 0 Then
Text1(6).Text = "0.0"
Else
Text1(6).Text = 2 * i
End If

Do While Not rst.EOF
If rst1.Fields(0) = rst.Fields(3) Then
Text1(5).Text = rst1.Fields(3)
Exit Do
End If
rst.MoveNext
Loop

Else
MsgBox ("Wrong Issue ID. Re-Enter")
Text2.SetFocus

End If

rst.MoveFirst
'err:
'MsgBox ("Wrong Issue ID! Re-Enter!")
'Text1(1).SetFocus
End Sub

Private Sub Form_Load()
conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "DB1.mdb;Mode=read|write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Issue"
.CommandType = adCmdText
End With

With rst
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd
End With

With cmd1
.ActiveConnection = conn
.CommandText = "SELECT * from Stock"
.CommandType = adCmdText
End With

With rst1
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd1
End With

Text1(0).Text = Form5.Text1.Text

End Sub

