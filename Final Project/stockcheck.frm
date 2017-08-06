VERSION 5.00
Begin VB.Form stockcheck 
   BackColor       =   &H80000004&
   Caption         =   "Stock"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14325
   LinkTopic       =   "Form3"
   ScaleHeight     =   6000
   ScaleWidth      =   14325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Return"
      Height          =   492
      Left            =   6000
      TabIndex        =   23
      Top             =   4920
      Width           =   2052
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Updation"
      Height          =   495
      Left            =   7080
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Requisition"
      Height          =   495
      Left            =   4920
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   8
      Left            =   12600
      TabIndex        =   20
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   7
      Left            =   11160
      TabIndex        =   19
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   6
      Left            =   9600
      TabIndex        =   18
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   5
      Left            =   8040
      TabIndex        =   17
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   4
      Left            =   6600
      TabIndex        =   16
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   3
      Left            =   4920
      TabIndex        =   15
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   2
      Left            =   3480
      TabIndex        =   14
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   1
      Left            =   1920
      TabIndex        =   13
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Last Updated"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   12600
      TabIndex        =   12
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4920
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   6600
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Publisher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8040
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "ISSD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   9600
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "ISSN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   11160
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Stock_id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Stock Id"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "stockcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
On Error Resume Next

With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Stock"
.CommandType = adCmdText
End With

Dim i As Integer
Dim unit As Integer
Dim flag As Integer
i = 0
flag = 0
Do While Not rst.EOF
    If rst.Fields(0) = Text1.Text Then
        flag = 1
        unit = CInt(rst.Fields(1))
        
        For i = 0 To 8 Step 1
            Label3(i).Caption = rst.Fields(i)
        Next
        Exit Do
        Exit Sub
        
    End If
    rst.MoveNext
Loop
rst.MoveFirst

If flag = 0 Then
    MsgBox ("Wrong Stock id")
    Exit Sub
End If

If unit < 20 Then
    MsgBox ("Stock amount is less than required")
    
End If
Command2.Visible = True
Command3.Visible = True


End Sub


Private Sub Command2_Click()
conn.Close
Unload Me
req.Show
End Sub

Private Sub Command3_Click()
conn.Close
Unload Me
stockadd.Show

End Sub

Private Sub Command4_Click()
conn.Close
Unload Me
Form1.Show
End Sub

Private Sub Form_Load()


conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "DB1.mdb;Mode=read|write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Stock"
.CommandType = adCmdText
End With

With rst
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd
End With

End Sub


