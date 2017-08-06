VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADMINISTRATOR PAGE"
   ClientHeight    =   7275
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   14175
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   7332
      Left            =   0
      Picture         =   "main.frx":0000
      ScaleHeight     =   7275
      ScaleWidth      =   14115
      TabIndex        =   0
      Top             =   0
      Width           =   14172
      Begin VB.CommandButton Command5 
         BackColor       =   &H000000FF&
         Caption         =   "Stock Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5400
         Width           =   1932
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Modify member"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5400
         Width           =   2052
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "Delete member"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   3000
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   2052
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "New member"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   720
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   2052
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H000000FF&
         Caption         =   "Display Member Details"
         DisabledPicture =   "main.frx":1F36F
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   7920
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   2052
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H8000000D&
         Caption         =   "Logout"
         Height          =   732
         Left            =   12000
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1812
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H000000FF&
         Caption         =   "Display Ex_ Member Details"
         DisabledPicture =   "main.frx":34D76
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   12240
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   2052
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Document Library Management System"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   19.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   12000
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Private Sub Command1_Click()
Form1.Hide
Form2.Show
Form2.Text1.Text = ""
Form2.Text2.Text = ""
Form2.Text3.Text = ""
Form2.Text4.Text = ""
Form2.Text5.Text = ""
End Sub



Private Sub Command2_Click()
Dim cmd1 As New ADODB.Command

d = InputBox("Enter the member code to delete the record", "Delete")
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
   App.Path & "\" & "DB1.mdb;Mode=Read|Write"
    conn.CursorLocation = adUseClient
    conn.Open

With cmd
.ActiveConnection = conn
  .CommandText = "select * from Issue where Mem_id='" + d + "'"
.CommandType = adCmdText
  End With

With rst
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd
    
    
End With
Do While Not rst.EOF
c = c + 1
rst.MoveNext
Loop
If c <> 0 Then
MsgBox "The member has not returned the book", vbInformation, "Library"
rst.Close
conn.Close
Else
cmd.CommandText = "select * from Member where Mem_id='" + d + "'"
rst.Close
rst.Open cmd
With cmd1

.ActiveConnection = conn
.CommandText = "Insert into Ex_member values('" & rst.Fields(0) & "','" & rst.Fields(1) & "','" & rst.Fields(2) & "','" & rst.Fields(3) & "','" & rst.Fields(4) & "','" & CDate(rst.Fields(5)) & "','" & CDate(11 / 4 / 5) & "','" & Abc & "')"
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data
End With


With rst
    
    .Delete
    .Requery
End With
rst.Close
MsgBox ("The member is deleted successfully")
End If

End Sub


Private Sub Command3_Click()
Unload Me
Form10.Show
End Sub

Private Sub Command4_Click()
Unload Me
Form6.Show

End Sub

Private Sub Form1_Load()

conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
   App.Path & "\" & "DB1.mdb;Mode=Read|Write"
    conn.CursorLocation = adUseClient
    conn.Open

With cmd
.ActiveConnection = conn
  .CommandText = "SELECT * From Member"
.CommandType = adCmdText
  End With

With rst
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd
    
    
End With


End Sub


Private Sub Command5_Click()
Unload Me
stockcheck.Show

End Sub

Private Sub Command6_Click()
Unload Me
Form4.Show
Load Form4


End Sub

Private Sub Command7_Click()
Unload Me
Form11.Show

End Sub

