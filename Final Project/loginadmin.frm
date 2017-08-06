VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H000000FF&
   Caption         =   "Admin's Login Form"
   ClientHeight    =   6810
   ClientLeft      =   105
   ClientTop       =   450
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
   ScaleHeight     =   6810
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   492
      Left            =   10800
      TabIndex        =   7
      Top             =   6240
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2772
      Left            =   7680
      TabIndex        =   1
      Top             =   3240
      Width           =   4452
      Begin VB.CommandButton Command1 
         Caption         =   "Sign In"
         Height          =   492
         Left            =   2640
         TabIndex        =   6
         Top             =   2040
         Width           =   1572
      End
      Begin VB.TextBox Text2 
         Height          =   348
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1200
         Width           =   2412
      End
      Begin VB.TextBox Text1 
         Height          =   348
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   2412
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Password"
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Username"
         Height          =   252
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1212
      End
   End
   Begin VB.Image Image1 
      Height          =   6090
      Left            =   0
      Picture         =   "loginadmin.frx":0000
      Top             =   2040
      Width           =   9135
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "Admin" And Text2.Text = "12345" Then
Form1.Show
'Form7.Show
Unload Me
Else
MsgBox "Admin Id or password incorrect"
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form4.Show

End Sub

