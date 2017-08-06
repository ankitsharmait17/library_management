VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form10 
   Caption         =   "Modification"
   ClientHeight    =   6936
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12288
   LinkTopic       =   "Form10"
   ScaleHeight     =   6936
   ScaleWidth      =   12288
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1572
   End
   Begin VB.CommandButton nxt 
      BackColor       =   &H8000000D&
      Caption         =   "Next"
      Height          =   840
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton prev 
      BackColor       =   &H8000000D&
      Caption         =   "Prev"
      Height          =   870
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1680
   End
   Begin VB.TextBox deptid 
      BackColor       =   &H8000000D&
      Height          =   732
      Left            =   3960
      TabIndex        =   4
      Top             =   5880
      Width           =   3768
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000D&
      DataMember      =   "dept id,dept name"
      Height          =   288
      Left            =   9480
      TabIndex        =   3
      Text            =   "Select the field"
      Top             =   1320
      Visible         =   0   'False
      Width           =   3372
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000D&
      Height          =   372
      Left            =   9480
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   3372
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   372
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   3492
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "Return"
      Height          =   732
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1932
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form10.frx":0000
      Height          =   3012
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   9012
      _ExtentX        =   15896
      _ExtentY        =   5313
      _Version        =   393216
      BackColor       =   -2147483635
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   612
      Left            =   11520
      Top             =   4680
      Visible         =   0   'False
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   1080
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\AM\Desktop\lib management\DB1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\AM\Desktop\lib management\DB1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *  from Member"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Document Library Management System"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12000
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Enter the Member id whose recored is to be modified"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   480
      TabIndex        =   9
      Top             =   5880
      Width           =   3132
   End
   Begin VB.Image Image1 
      Height          =   8196
      Left            =   0
      Picture         =   "Form10.frx":0015
      Top             =   0
      Width           =   12288
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Combo1_Click()
MsgBox ("enter the new entry")
Text2.Visible = True
Text2.SetFocus
End Sub

Private Sub Command1_Click()

With cmd
.ActiveConnection = conn
If Combo1.Text = "M_name" Then
.CommandText = "update Member set M_name = '" & Text2.Text & "' where Mem_id='" & deptid.Text & "'"

End If
If Combo1.Text = "M_addr" Then
.CommandText = "update Member set M_addr = '" & Text2.Text & "' where Mem_id='" & deptid.Text & "'"
End If

If Combo1.Text = "M_cont" Then
.CommandText = "update Member set M_cont = '" & Text2.Text & "' where Mem_id='" & deptid.Text & "'"
End If
If Combo1.Text = "M_prof" Then
.CommandText = "update Member set M_prof = '" & Text2.Text & "' where Mem_id='" & deptid.Text & "'"
End If
If Combo1.Text = "Mem_id" Then
.CommandText = "update Member set Mem_id = '" & Text2.Text & "' where Mem_id='" & deptid.Text & "'"
End If
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data
MsgBox ("record updated successfully")
.CommandText = "select * from Member"
rst.Close
rst.Open cmd
Adodc1.Refresh
DataGrid1.Refresh
   ' rst.AbsolutePosition = rwid
    'Adodc1.Recordset.AbsolutePosition = rwid
End With
deptid.Text = ""
Combo1.Text = ""
Text2.Text = ""

End Sub

Private Sub Command2_Click()
If Len(deptid.Text) = 0 Then
    MsgBox ("Member id is mandatory for update")
    deptid.SetFocus
End If
cmd.CommandText = "select * from Member where Mem_id='" + deptid.Text + "'"
rst.Close
rst.Open cmd
Do While Not rst.EOF
c = c + 1
rst.MoveNext
Loop
If c <> 0 Then
MsgBox ("Select the feild to be updated")
Combo1.Visible = True
Else
deptid.Text = ""
MsgBox "Incorrect Member Id", vbExclamation, "ERROR"
End If

End Sub

Private Sub Command3_Click()
Unload Me

Form1.Show

End Sub

Private Sub Form_Load()
Combo1.AddItem ("Mem_id")
Combo1.AddItem ("M_name")
Combo1.AddItem ("M_addr")
Combo1.AddItem ("M_prof")
Combo1.AddItem ("M_cont")



conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "DB1.mdb;Mode=read|write"
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

Private Sub nxt_Click()
On Error GoTo er

If Not rst.EOF Then
    rst.MoveNext
    Adodc1.Recordset.MoveNext
   ' Call showrcd
End If

Exit Sub

er:
MsgBox "Reached end of file,moving to first"
rst.MoveFirst
Adodc1.Recordset.MoveFirst

'Call showrcd
End Sub

Private Sub prev_Click()
On Error GoTo er

If Not rst.EOF Then
    rst.MovePrevious
    Adodc1.Recordset.MovePrevious
    'Call showrcd
End If

Exit Sub

er:
MsgBox ("Reached BOF,moving to last")
rst.MoveLast
Adodc1.Recordset.MoveLast

'Call showrcd

End Sub

Private Sub Text2_Change()
Command1.Enabled = True


End Sub


