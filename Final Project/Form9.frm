VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   6696
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9984
   LinkTopic       =   "Form9"
   ScaleHeight     =   6696
   ScaleWidth      =   9984
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
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
      TabIndex        =   15
      Top             =   5160
      Width           =   4332
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   3480
      Width           =   4332
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   2760
      Width           =   4332
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   4332
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   4332
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   4332
   End
   Begin VB.CommandButton ret 
      Caption         =   "Back"
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton mem 
      Caption         =   "Return Book"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
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
      TabIndex        =   0
      Top             =   4200
      Width           =   4332
   End
   Begin VB.Label Label8 
      Caption         =   "Fine"
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
      TabIndex        =   16
      Top             =   5160
      Width           =   2052
   End
   Begin VB.Label Label6 
      Caption         =   "BOOK RETURN"
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
      TabIndex        =   14
      Top             =   0
      Width           =   3732
   End
   Begin VB.Label Label5 
      Caption         =   "Submission Date"
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
      TabIndex        =   13
      Top             =   4200
      Width           =   2052
   End
   Begin VB.Label Label4 
      Caption         =   "Return Date"
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
      TabIndex        =   12
      Top             =   3600
      Width           =   2652
   End
   Begin VB.Label Label3 
      Caption         =   "Mem_id"
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
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Issue_date"
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
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Issue_id"
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
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Stock_id"
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
      TabIndex        =   8
      Top             =   2880
      Width           =   2292
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
