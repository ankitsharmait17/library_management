VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H000000FF&
   Caption         =   "User Login Form"
   ClientHeight    =   6804
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12660
   FillColor       =   &H00FF80FF&
   BeginProperty Font 
      Name            =   "Rockwell"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   6804
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   492
      Left            =   10800
      TabIndex        =   5
      Top             =   6000
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2412
      Left            =   7560
      TabIndex        =   1
      Top             =   3240
      Width           =   4452
      Begin VB.CommandButton Command1 
         Caption         =   "Sign in"
         Height          =   492
         Left            =   2760
         TabIndex        =   4
         Top             =   1800
         Width           =   1572
      End
      Begin VB.TextBox Text1 
         Height          =   348
         Left            =   1800
         TabIndex        =   3
         Top             =   960
         Width           =   2412
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Enter Member ID No."
         Height          =   1092
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1212
      End
   End
   Begin VB.Image Image1 
      Height          =   4872
      Left            =   0
      Picture         =   "login.frx":0000
      Top             =   2040
      Width           =   7308
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Welcome to Document Library Management System"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2052
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7320
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset



Private Sub Command2_Click()
Unload Me
Form4.Show
Load Form4


End Sub

Private Sub Form_Load()

conn.ConnectionString = "Provider=Microsoft.jet.OLEDB.4.0;Data source =" & App.Path & "\" & "DB1.mdb;Mode=read|write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Member"
.CommandType = adCmdText
End With

With rst
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd
End With
End Sub

Private Sub Command1_Click()
'On Error GoTo errorhandle
cmd.CommandText = "select * from Member where Mem_id='" + Text1.Text + "'"
rst.Close
rst.Open cmd
Do While Not rst.EOF
c = c + 1
rst.MoveNext
Loop
If c <> 0 Then
MsgBox "Login Successful", vbExclamation, "LOGIN"
conn.Close


Optn.Show
Else
Text1.Text = ""
MsgBox "Incorrect Member Id", vbExclamation, "LOGIN UNSUCCESSFUL"
End If
End Sub

