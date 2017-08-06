VERSION 5.00
Begin VB.Form Issue 
   Caption         =   "Document Issue"
   ClientHeight    =   8544
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   14076
   LinkTopic       =   "Form1"
   ScaleHeight     =   8544
   ScaleWidth      =   14076
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text10 
      Height          =   615
      Left            =   11280
      TabIndex        =   20
      Text            =   "Text10"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   10680
      TabIndex        =   19
      Text            =   "Issued"
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Text            =   "000.0001"
      Top             =   960
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   5880
      TabIndex        =   17
      Text            =   "Text7"
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ISSUE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   16
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      Height          =   405
      Left            =   8400
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      Height          =   405
      Left            =   8400
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   405
      Left            =   8400
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Top             =   5760
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   2352
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   405
      Left            =   8400
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CONFIRM"
      Enabled         =   0   'False
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   4440
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Expected Submission Date"
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Issue Date"
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Enter Quantity"
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Member ID"
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   3960
      Width           =   3135
   End
End
Attribute VB_Name = "Issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim rst1 As New ADODB.Recordset
Private Sub Combo1_Click()
Label2.Caption = "Select " & Combo1.Text & " From Below"
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

Dim temp As String
List1.Clear
rst.MoveFirst

Do
temp = rst.Fields(2)
If temp = Combo1.Text Then
List1.AddItem (rst.Fields(3))
End If
rst.MoveNext
If rst.EOF Then
    Exit Do
End If
Loop
conn.Close


End Sub

Private Sub Command1_Click()
'On Error GoTo er
Dim i As Integer
conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "DB1.mdb;Mode=read|write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "Select * from Issue"
.CommandType = adCmdText
End With
With rst
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd
End With

rst.MoveLast
i = Val(rst.Fields(0))
i = i + 1
Text10.Text = i
rst.AddNew
rst.Fields(0) = Text10.Text
rst.Fields(1) = CDate(Text5.Text)
rst.Fields(2) = Text3.Text
rst.Fields(3) = Text7.Text
rst.Fields(4) = CDate(Text6.Text)
rst.Fields(5) = DateValue(Now)
rst.Fields(6) = Text8.Text
rst.Fields(7) = Text9.Text

rst.Update
'With cmd
'.ActiveConnection = conn
'.CommandText = "Insert Into Issue values(Text10.Text,CDate(Text5.Text),Text3.Text,Text7.Text,CDate(Text6.Text),DateValue(Now),Text8.Text,Text9.Text)"
'.CommandType = adCmdText
'conn.BeginTrans 'to insert a new row
'.Execute 'to insert the data
'conn.CommitTrans 'to save the data
'End With

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

rst1.MoveFirst

Do
If rst1.Fields(3) = Text1.Text Then
rst1.Fields(1) = Val(rst1.Fields(1)) - Val(Text4.Text)
Exit Do
End If
rst1.MoveNext
Loop
rst1.Update
conn.Close
MsgBox ("Issued")
MsgBox ("Your Issue ID is " & Text10.Text & ". Please remember this ID. It will be required at the time of return.")
Unload Me

'MsgBox (Adodc1.Recordset!Table2)

Exit Sub

'er:
'MsgBox ("No Record Found")

End Sub

Private Sub Command2_Click()

conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "DB1.mdb;Mode=read|write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Stock where title='" & Text1.Text & "'"
.CommandType = adCmdText
End With

With rst
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd
End With

Text3.Text = Form5.Text1.Text

Text7.Text = rst.Fields(0)

Label3.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True

Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True

Text4.Enabled = True

conn.Close

End Sub

Private Sub Form_Load()
Combo1.AddItem "Book"
Combo1.AddItem "CD"
Combo1.AddItem "DVD"
Combo1.AddItem "Magazine"
Combo1.AddItem "Journal"
Combo1.Text = "Book"


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

Call showrcd

Do
List1.AddItem (rst.Fields(3))
rst.MoveNext
If rst.EOF Then
    Exit Do
End If
Loop

rst.MoveFirst
conn.Close

End Sub

Function showrcd()
'deptid.Text = rst.Fields(0) 'when not known
'deptname.Text = rst!stock 'when we know the feild name
'rwid = rst.AbsolutePosition
End Function

Private Sub List1_Click()

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

Label1.Caption = "Number of " & Combo1.Text & "s Available:"
Label4.Caption = " " & Combo1.Text & " Selected:"
Text1.Text = List1.Text

Do
If rst.Fields(3) = Text1.Text Then
Text2.Text = rst.Fields(1)
Exit Do
End If
rst.MoveNext
Loop
If Text2.Text > 0 Then
Command2.Enabled = True
End If

conn.Close
End Sub




Private Sub Text4_LostFocus()

If Val(Text4.Text) > Val(Text2.Text) Then
MsgBox ("Entered Quatity is greater than available quantity. Enter Again!")
Text4.SetFocus
Exit Sub
End If

Text5.Text = DateValue(Now)
Text6.Text = DateValue(DateAdd("d", 21, Now))
Command1.Enabled = True
Command1.SetFocus
End Sub
