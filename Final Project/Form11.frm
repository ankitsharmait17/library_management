VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form11 
   Caption         =   "Ex-Member Details"
   ClientHeight    =   6144
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11484
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   6144
   ScaleWidth      =   11484
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   732
      Left            =   4200
      TabIndex        =   3
      Top             =   5160
      Width           =   2532
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Text            =   "Enter the ID"
      Top             =   5280
      Width           =   3612
   End
   Begin VB.CommandButton Command2 
      Caption         =   "All Records"
      Height          =   492
      Left            =   8160
      TabIndex        =   1
      Top             =   0
      Width           =   2052
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Return"
      Height          =   492
      Left            =   8160
      TabIndex        =   0
      Top             =   5280
      Width           =   2052
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   492
      Left            =   960
      Top             =   0
      Visible         =   0   'False
      Width           =   6492
      _ExtentX        =   11451
      _ExtentY        =   868
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "SELECT *  from Ex_member"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form11.frx":1302
      Height          =   4332
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   11052
      _ExtentX        =   19495
      _ExtentY        =   7641
      _Version        =   393216
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
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
   App.Path & "\" & "DB1.mdb;Mode=Read|Write"
    conn.CursorLocation = adUseClient
    conn.Open

With cmd
.ActiveConnection = conn
  .CommandText = "SELECT * From Ex_member where Mem_id='" & Text1.Text & "'"
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
If c = 0 Then
MsgBox ("wrong member id")
Else
Adodc1.Refresh
DataGrid1.Refresh
Adodc1.RecordSource = "SELECT * From Ex_member where Mem_id='" & Text1.Text & "'"
conn.Close
End If
End Sub

Private Sub Command2_Click()
Adodc1.Refresh
DataGrid1.Refresh
Adodc1.RecordSource = "SELECT * From Ex_member"

End Sub

Private Sub Command3_Click()
Unload Me
Form1.Show

End Sub

