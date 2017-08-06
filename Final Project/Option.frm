VERSION 5.00
Begin VB.Form Optn 
   Caption         =   "User Options"
   ClientHeight    =   4716
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   9432
   LinkTopic       =   "Form2"
   ScaleHeight     =   4716
   ScaleWidth      =   9432
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Logout"
      Height          =   975
      Left            =   5160
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Issue  Details"
      Height          =   975
      Left            =   1440
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Document Return"
      Height          =   975
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Document Issue"
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Optn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Issue.Show
End Sub

Private Sub Command2_Click()
Retur.Show

End Sub

Private Sub Command3_Click()
IssueD.Show
End Sub

Private Sub Command4_Click()
Unload Me
Form4.Show
End Sub
