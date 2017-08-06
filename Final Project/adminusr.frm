VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000D&
   ClientHeight    =   6660
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12516
   LinkTopic       =   "Form4"
   ScaleHeight     =   6660
   ScaleWidth      =   12516
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "USER"
      Height          =   732
      Left            =   6240
      TabIndex        =   1
      Top             =   4320
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADMIN"
      Height          =   732
      Left            =   6240
      TabIndex        =   0
      Top             =   2280
      Width           =   1932
   End
   Begin VB.Image Image2 
      Height          =   1224
      Left            =   3960
      Picture         =   "adminusr.frx":0000
      Top             =   4080
      Width           =   1224
   End
   Begin VB.Image Image1 
      Height          =   1224
      Left            =   3960
      Picture         =   "adminusr.frx":0689
      Top             =   2040
      Width           =   1224
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form3.Show




End Sub

Private Sub Command2_Click()
Unload Me
Form5.Show

End Sub
