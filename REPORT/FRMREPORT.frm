VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMREPORT 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORT VIEWING FORM"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   5520
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF00FF&
      Caption         =   "ACCES MASTER RECORD REPORT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   600
      TabIndex        =   18
      Top             =   5880
      Width           =   5775
      Begin VB.CommandButton Command5 
         Caption         =   "VIEW ALL RECORD "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   24
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "SEARCH RECORDS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   375
         Left            =   840
         TabIndex        =   20
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   46137345
         CurrentDate     =   41662
      End
      Begin MSComCtl2.DTPicker DTPicker8 
         Height          =   375
         Left            =   3120
         TabIndex        =   21
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   46137345
         CurrentDate     =   41662
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "FROM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000000FF&
      Caption         =   "RECORD UPDATE REPORT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   600
      TabIndex        =   8
      Top             =   4080
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "SEARCH RECORDS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   9
         Top             =   1080
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   46137345
         CurrentDate     =   41662
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   46137345
         CurrentDate     =   41662
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "FROM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "GOODS RECEIVE REPORT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "SEARCH RECORDS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   46137345
         CurrentDate     =   41662
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   46137345
         CurrentDate     =   41662
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FROM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "SALES REPORT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "SEARCH RECORDS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   1080
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   46137345
         CurrentDate     =   41662
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   46137345
         CurrentDate     =   41662
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "FROM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   10
      Height          =   7695
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "FRMREPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

