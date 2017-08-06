VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form stockadd 
   Caption         =   "Stock Updation"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13395
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   495
      Left            =   12120
      TabIndex        =   25
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return"
      Height          =   612
      Left            =   8160
      TabIndex        =   24
      Top             =   6960
      Width           =   1692
   End
   Begin VB.TextBox Text9 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd""/""MM""/""yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton nw 
      Caption         =   "New"
      Height          =   495
      Left            =   9120
      TabIndex        =   21
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   10560
      TabIndex        =   20
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton ins 
      Caption         =   "Insert"
      Height          =   495
      Left            =   7680
      TabIndex        =   19
      Top             =   6240
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "stockadd.frx":0000
      Height          =   5055
      Left            =   4320
      TabIndex        =   18
      Top             =   480
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8916
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Height          =   492
      Left            =   7920
      Top             =   7680
      Visible         =   0   'False
      Width           =   2052
      _ExtentX        =   3625
      _ExtentY        =   873
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Stock"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton nxt 
      Caption         =   "Next"
      Height          =   495
      Left            =   6120
      TabIndex        =   17
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton prev 
      Caption         =   "Prev"
      Height          =   495
      Left            =   4560
      TabIndex        =   16
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Date of Updation"
      Height          =   495
      Left            =   840
      TabIndex        =   22
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Id_desc"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Id_no"
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Stock_desc"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Stock Type"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "title"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Author"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "publisher"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Stock_id"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "stockadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim rwid


Private Sub Command1_Click()
conn.Close
Unload Me
stockcheck.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next

Dim d As Date
If Len(Text1.Text) = 0 Then
    MsgBox ("Record required for update")
    Text1.SetFocus
    Exit Sub
End If
With cmd
.ActiveConnection = conn
.CommandText = "update Stock set Stock_desc='" & Text2.Text & "',Last_upd='" & CDate(Text9.Text) & " ' where Stock_id='" & Text1.Text & "'"
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data

MsgBox ("record updated successfully")
.CommandText = "select * from Stock"
rst.Close
    rst.Open cmd
    
    Adodc1.Refresh
    DataGrid1.Refresh
    rst.AbsolutePosition = rwid
    Adodc1.Recordset.AbsolutePosition = rwid
End With

End Sub

Private Sub Command3_Click()
With cmd
.ActiveConnection = conn
.CommandText = "delete from Stock where Stock_id='" & Text1.Text & "'"
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data

MsgBox ("record deleted successfully")
.CommandText = "select * from Stock"
rst.Close
    rst.Open cmd
    
    Adodc1.Refresh
    DataGrid1.Refresh
End With
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
Function showrcd()
Text1.Text = rst.Fields(0)
Text2.Text = rst.Fields(1)
Text3.Text = rst.Fields(2)
Text4.Text = rst.Fields(3)
Text5.Text = rst.Fields(4)
Text6.Text = rst.Fields(5)
Text7.Text = rst.Fields(6)
Text8.Text = rst.Fields(7)
Text9.Text = CDate(rst.Fields(8))

rwid = rst.AbsolutePosition
End Function

Private Sub ins_Click()

If Len(Text1.Text) = 0 Then
    MsgBox ("Record required for insert")
    Text1.SetFocus
    Exit Sub
End If

With cmd
.ActiveConnection = conn
.CommandText = "Insert into Stock values('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & CDate(Text9.Text) & "')"
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data

MsgBox ("data inserted successfully")
.CommandText = "select * from Stock"
rst.Close
rst.Open cmd
Adodc1.Refresh
DataGrid1.Refresh

rst.MoveLast
Adodc1.Recordset.MoveLast

End With
End Sub

Private Sub nw_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""

End Sub

Private Sub nxt_Click()
On Error GoTo er

If Not rst.EOF Then
    rst.MoveNext
    Adodc1.Recordset.MoveNext
    Call showrcd
End If

Exit Sub

er:
MsgBox "Reached end of file,moving to first"
rst.MoveFirst
Adodc1.Recordset.MoveFirst

Call showrcd

End Sub

Private Sub prev_Click()
On Error GoTo er

If Not rst.EOF Then
    rst.MovePrevious
    Adodc1.Recordset.MovePrevious
    Call showrcd
End If

Exit Sub

er:
MsgBox ("Reached BOF,moving to last")
rst.MoveLast
Adodc1.Recordset.MoveLast

Call showrcd
End Sub

