VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   6336
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12288
   LinkTopic       =   "Form7"
   ScaleHeight     =   6336
   ScaleWidth      =   12288
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form7.frx":0000
      Height          =   2652
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   9372
      _ExtentX        =   16531
      _ExtentY        =   4678
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
   Begin VB.CommandButton Command3 
      Caption         =   "Issue Details"
      Height          =   732
      Left            =   2760
      TabIndex        =   2
      Top             =   4920
      Width           =   2052
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Return Book"
      Height          =   732
      Left            =   8760
      TabIndex        =   1
      Top             =   4920
      Width           =   2172
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Issue New Book"
      Height          =   732
      Left            =   5880
      TabIndex        =   0
      Top             =   4920
      Width           =   1932
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
Load Form8
End Sub

Private Sub Command2_Click()
Unload Me
Load Form9


End Sub

Private Sub Command3_Click()
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Issue where Mem_id = '" & Form5.Text1.Text & "'"
.CommandType = adCmdText
End With



End Sub

Private Sub Form_Load()

conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "DB1.mdb;Mode=read|write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Issue"
.CommandType = adCmdText
End With

With rst
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd
End With
End Sub

