VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6048
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11784
   LinkTopic       =   "Form2"
   ScaleHeight     =   6048
   ScaleWidth      =   11784
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "M/d/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   4200
      Width           =   4332
   End
   Begin VB.CommandButton mem 
      Caption         =   "&Add Member"
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton ret 
      Caption         =   "&Return"
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   4332
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1560
      Width           =   4332
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2160
      Width           =   4332
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   4332
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   3480
      Width           =   4332
   End
   Begin VB.Label Label7 
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   372
      Left            =   0
      TabIndex        =   13
      Top             =   2880
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Profession"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      TabIndex        =   7
      Top             =   3600
      Width           =   2652
   End
   Begin VB.Label Label5 
      Caption         =   "Registration Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      TabIndex        =   6
      Top             =   4200
      Width           =   2052
   End
   Begin VB.Label Label6 
      Caption         =   "Member Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1800
      TabIndex        =   5
      Top             =   0
      Width           =   3732
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Private Sub Form_Load()

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

Private Sub mem_Click()
cmd.CommandText = "select * from Member"
rst.AddNew
rst.Fields(0) = Text1.Text
rst.Fields(1) = Text2.Text
rst.Fields(2) = Text3.Text
rst.Fields(3) = Text4.Text
rst.Fields(4) = Text5.Text
rst.Fields(5) = CDate(Text6.Text)
rst.Update
rst.Close
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub ret_Click()
Form2.Hide
Form1.Show

End Sub

Private Sub Text1_LostFocus()
On Error GoTo errorhandle
cmd.CommandText = "select * from Member where Mem_id='" + Text1.Text + "'"
rst.Close
rst.Open cmd
Do While Not rst.EOF
c = c + 1
rst.MoveNext
Loop
If c <> 0 Then
MsgBox "User id already exists", vbExclamation, "Duplicate"
Text1.Text = " "
Text1.SetFocus
Else
Text2.SetFocus
End If
Exit Sub
errorhandle:
MsgBox "Error occurred!Wrong Member Code", vbInformation, "Error"
End
End Sub

