VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMREPORT 
   BackColor       =   &H00FF0000&
   Caption         =   "REPORT VIEWING FORM"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12120
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   9120
      Visible         =   0   'False
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
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
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   1920
      TabIndex        =   14
      Top             =   240
      Width           =   8415
      Begin VB.CommandButton Command7 
         Caption         =   "SEARCH RECORDS BY DATE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   23
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5400
         TabIndex        =   22
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SEARCH RECORDS BY DATE AND GOODS TYPE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   15
         Top             =   1440
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   14.25
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
         Format          =   79888385
         CurrentDate     =   41662
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   3120
         TabIndex        =   17
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   14.25
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
         Format          =   79888385
         CurrentDate     =   41662
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GOODS TYPE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   39
         Top             =   360
         Width           =   2295
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
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   735
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
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      Caption         =   " GOODS RECEIVED REPORT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   1920
      TabIndex        =   8
      Top             =   2400
      Width           =   8415
      Begin VB.CommandButton Command16 
         Caption         =   "SEARCH RECORDS BY DATE AND GOODS TYPE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   29
         Top             =   1440
         Width           =   4815
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5400
         TabIndex        =   28
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SEARCH RECORDS "
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
         Left            =   720
         TabIndex        =   9
         Top             =   840
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   495
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   79888385
         CurrentDate     =   41662
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   495
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   79888385
         CurrentDate     =   41662
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GOODS TYPE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   40
         Top             =   360
         Width           =   2295
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
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   360
         Width           =   975
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
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
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
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   1920
      TabIndex        =   2
      Top             =   4680
      Width           =   8415
      Begin VB.CommandButton Command12 
         Caption         =   "SEARCH RECORDS BY DATE AND GOODS TYPE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   27
         Top             =   1440
         Width           =   5055
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5520
         TabIndex        =   26
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SEARCH RECORDS BY DATE "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   495
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   79888385
         CurrentDate     =   41662
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   495
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65535
         CalendarForeColor=   255
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16711680
         Format          =   79888385
         CurrentDate     =   41662
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GOODS TYPE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   41
         Top             =   360
         Width           =   2655
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
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   735
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
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H008080FF&
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
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   480
      TabIndex        =   0
      Top             =   6960
      Width           =   11415
      Begin VB.CommandButton Command9 
         Caption         =   "SEARCH"
         Height          =   495
         Left            =   10800
         TabIndex        =   38
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "LESS THAN"
         Height          =   375
         Left            =   10800
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "GREATER"
         Height          =   375
         Left            =   10800
         TabIndex        =   36
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   10800
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ALL POSITIVE STOCK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8640
         TabIndex        =   33
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton CMDZEROSTOCKS 
         Caption         =   "ALL NEGATIVE STOCKS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8640
         TabIndex        =   32
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command19 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12360
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12360
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "SEARCH BY GOODS LOCATION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   3135
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3360
         TabIndex        =   24
         Top             =   1800
         Width           =   4455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "SEARCH BY GOODS TYPE "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3360
         TabIndex        =   20
         Top             =   1080
         Width           =   4455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "VIEW ALL RECORDS IN MASTER "
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
         TabIndex        =   1
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER GOODS CODE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12240
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   9375
      Left            =   240
      Picture         =   "FRMREPORT.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   11655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000040&
      BorderWidth     =   10
      Height          =   9615
      Left            =   120
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "FRMREPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDZEROSTOCKS_Click()

Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

Dim REMAIN
REMAIN = 1


'query to filter the table

RS.Open "insert into [MASTERREPORT] SELECT Code ,Description,Location,GoodsType,Stock,Price,DateModify,UNIT FROM [Master] where Stock < " & REMAIN & " ORDER BY GoodsType,DESCRIPTION ASC", DB, 3, 3


'RS.Open "insert into [MASTERREPORT] SELECT Code ,Description,GoodsType,Location,Stock,Price,DateModify,UNIT FROM [Master] where Stock < @" & REMAIN & " ORDER BY GoodsType,DESCRIPTION ASC", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Master", acViewPreview
APPACCESS.Visible = True

'DELETE TABLE
RS12.Open "DELETE * FROM MASTERREPORT", DB, 3, 3


End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table
RS.Open "insert into [SOLDREPORT] SELECT Code, Description, Location, Stock, Price, CustName, TDate,Quantity,Discount,GoodsType,Amount,MEASURE FROM [Sold] where TDate between #" & Str(DTPicker1.Value) & "# and #" & Str(DTPicker2.Value) & "# AND GoodsType = '" & Combo2.Text & "'ORDER BY ID", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Sold", acViewPreview
APPACCESS.Visible = True
'DELETE TABLE
RS12.Open "DELETE * FROM SOLDREPORT ", DB, 3, 3
End Sub


Private Sub Command11_Click()
Dim DB333 As New ADODB.CONNECTION
Dim RS331 As New ADODB.Recordset
Dim SQL As String
Dim RS333 As New ADODB.Recordset
DB333.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table
RS333.Open "insert into [MASTERMODIFYREPORT] SELECT OldCode, NewCode,OldDescription,NewDescription, OldLocation,NewLocation, OldStock,NewStock, OldPrice,NewPrice,Date_Modify,OLDMEASURE,NEWMEASURE FROM [MasterModify] where GoodsType = '" & Combo5.Text & "' ORDER BY ID", DB333, 3, 3

'RS.Open " where Date_Modify between #" & Str(DTPicker5.Value) & "# and #" & Str(DTPicker6.Value) & "# ORDER BY ID", DB, 3, 3

'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "MasterModify", acViewPreview
APPACCESS.Visible = True

'DELETE THE TABLE
RS331.Open "DELETE * FROM MASTERMODIFYREPORT", DB333, 3, 3
End Sub

Private Sub Command12_Click()

Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table

RS.Open "insert into [MASTERMODIFYREPORT] SELECT OldCode, NewCode,OldDescription,NewDescription, OldLocation,NewLocation, OldStock,NewStock, OldPrice,NewPrice,Date_Modify,OLDMEASURE,NEWMEASURE FROM [MasterModify] where Date_Modify between #" & Str(DTPicker5.Value) & "# and #" & Str(DTPicker6.Value) & "# AND GoodsType = '" & Combo5.Text & "'ORDER BY ID", DB, 3, 3

'RS.Open "insert into [SOLDREPORT] SELECT Code, Description, Location, Stock, Price, CustName, TDate,Quantity,Discount,GoodsType,Amount,MEASURE FROM [Sold] where TDate between #" & Str(DTPicker1.Value) & "# and #" & Str(DTPicker2.Value) & "# AND GoodsType = '" & Combo2.Text & "'ORDER BY ID", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "MasterModify", acViewPreview
APPACCESS.Visible = True
'DELETE TABLE
RS12.Open "DELETE * FROM MASTERMODIFYREPORT ", DB, 3, 3

End Sub

Private Sub Command15_Click()

Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table


'RS.Open "insert into [RECEIVEDREPORT] SELECT Code, Description, Location, OldStock, Price,Qsupply,Amount, RDate,New_Stock,SuppName,SuppContact,SuppComp,MEASURE FROM [Received] where RDate between #" & Str(DTPicker3.Value) & "# and #" & Str(DTPicker4.Value) & "# ORDER BY ID ", DB, 3, 3


RS.Open "insert into [RECEIVEDREPORT] SELECT Code, Description, Location, OldStock, Price,Qsupply,Amount, RDate,New_Stock,SuppName,SuppContact,SuppComp,MEASURE FROM [Received] where GoodsType= '" & Combo7.Text & "' ORDER BY ID ", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Received", acViewPreview
APPACCESS.Visible = True
'DELETE TABLE
RS12.Open "DELETE * FROM RECEIVEDREPORT", DB, 3, 3

End Sub

Private Sub Command16_Click()


Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table

'RS.Open "insert into [RECEIVEDREPORT]  ORDER BY ID ", DB, 3, 3

RS.Open "insert into [RECEIVEDREPORT]  SELECT Code, Description, Location, OldStock, OldPrice,Price,Qsupply,Amount, RDate,New_Stock,SuppName,SuppContact,SuppComp,MEASURE FROM [Received] where RDate between #" & Str(DTPicker3.Value) & "# and #" & Str(DTPicker4.Value) & "# AND GoodsType = '" & Combo7.Text & "'ORDER BY ID", DB, 3, 3

'RS.Open "insert into [SOLDREPORT] SELECT Code, Description, Location, Stock, Price, CustName, TDate,Quantity,Discount,GoodsType,Amount,MEASURE FROM [Sold] where TDate between #" & Str(DTPicker1.Value) & "# and #" & Str(DTPicker2.Value) & "# AND GoodsType = '" & Combo2.Text & "'ORDER BY ID", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Received", acViewPreview
APPACCESS.Visible = True
'DELETE TABLE
RS12.Open "DELETE * FROM RECEIVEDREPORT ", DB, 3, 3


End Sub

Private Sub Command19_Click()
If Text1.Text = "" Then
MsgBox "enter a Valid goods code to check for history ....in these box", vbInformation, "EMPTY / NO GOODS CODE ENTERED"
Text1.SetFocus
Exit Sub
Else
Dim DB211 As New ADODB.CONNECTION
Dim RS211 As New ADODB.Recordset
Dim RS212 As New ADODB.Recordset
Dim RS213 As New ADODB.Recordset
Dim RS214 As New ADODB.Recordset
''''master''''''''
Dim M1, M2, M3, M4, M5, M6

'''''' OPEN THE DATABASE MARKET''''''''''''
DB211.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

''''''SEARCH MASTER TABLE ''''''''

If RS211.State = adStateOpen Then RS211.Close
'RS211.Open "insert into [HISTORY] SELECT Description, DateReg, Stock,unit, Price,New FROM [Master] where code = '" & Trim(Text1.Text) & "'", DB211, 3, 3

RS211.Open "SELECT* FROM [Master] where code = '" & Trim(Text1.Text) & "'", DB211, adOpenKeyset, adLockReadOnly
If RS211.EOF Then
Call checkreceived
Else
RS211.MoveFirst
Do Until RS211.EOF
M1 = RS211!Description
M2 = RS211!DateReg
M3 = RS211!Stock
M4 = RS211!unit
M5 = RS211!Price
M6 = RS211!New

Dim RS400 As New ADODB.Recordset
RS400.Open "SELECT* FROM [HISTORY]", DB211, 3, 3
RS400.AddNew
RS400!Description = M1
RS400!HDate = M2
RS400!Stock = M3
RS400!MEASURE = M4
RS400!Price = M5
RS400!Operation = M6
RS400.Update
RS400.Close
Set RS400 = Nothing
RS211.MoveNext
Loop
RS211.Close
Set RS211 = Nothing
Call checkreceived
End If
End If
End Sub
Public Sub checkreceived()
Dim DB711 As New ADODB.CONNECTION
Dim RS711 As New ADODB.Recordset
Dim RS712 As New ADODB.Recordset

''''received''''''
Dim M1A, M2A, M3A, M4A, M5A, M6A
If DB711.State = adStateOpen Then DB711.Close
DB711.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

If RS711.State = adStateOpen Then RS211.Close
RS711.Open "SELECT * FROM [Received] where code = '" & Trim(Text1.Text) & "'", DB711, adOpenKeyset, adLockReadOnly
If RS711.EOF Then
Call checksold
Else
RS711.MoveFirst
Do Until RS711.EOF
MIA = RS711!Description
M2A = RS711!RDate
M3A = RS711!New_Stock
M4A = RS711!MEASURE
M5A = RS711!Price
M6A = RS711!RECEIVED

'Dim RS400 As New ADODB.Recordset
If RS712.State = adStateOpen Then RS712.Close
RS712.Open "SELECT* FROM [HISTORY]", DB711, 3, 3
RS712.AddNew
RS712!Description = M1A
RS712!HDate = M2A
RS712!Stock = M3A
RS712!MEASURE = M4A
RS712!Price = M5A
RS712!Operation = M6A
RS712.Update
RS712.Close
Set RS712 = Nothing
RS711.MoveNext
Loop
RS711.Close
Set RS711 = Nothing
Call checksold
End If
End Sub
Public Sub checksold()

Dim DB713 As New ADODB.CONNECTION
Dim RS713 As New ADODB.Recordset
Dim RS714 As New ADODB.Recordset

''''''''SOLD'''''''
Dim M1AA, M2AA, M3AA, M4AA, M5AA, M6AA
If DB713.State = adStateOpen Then DB713.Close
DB713.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

If RS713.State = adStateOpen Then RS713.Close
RS713.Open "SELECT * FROM [sold] where code = '" & Trim(Text1.Text) & "'", DB713, adOpenKeyset, adLockReadOnly
If RS713.EOF Then
Call checkmodify
Else
RS713.MoveFirst
Do Until RS713.EOF
M1AA = RS713!Description
M2AA = RS713!TDate
M3AA = RS713!Quantity
M4AA = RS713!MEASURE
M5AA = RS713!Price
M6AA = RS713!sold

'Dim RS400 As New ADODB.Recordset
If RS714.State = adStateOpen Then RS714.Close
RS714.Open "SELECT * FROM [HISTORY]", DB713, 3, 3
RS714.AddNew
RS714!Description = M1AA
RS714!HDate = M2AA
RS714!Stock = M3AA
RS714!MEASURE = M4AA
RS714!Price = M5AA
RS714!Operation = M6AA
RS714.Update
RS714.Close
Set RS714 = Nothing
RS713.MoveNext
Loop
RS713.Close
Set RS713 = Nothing
Call checkmodify
End If
End Sub
Public Sub checkmodify()

Dim DB715 As New ADODB.CONNECTION
Dim RS715 As New ADODB.Recordset
Dim RS716 As New ADODB.Recordset

''''''MODIFY '''''''''
Dim M1AAA, M2AAA, M3AAA, M4AAA, M5AAA, M6AAA

If DB715.State = adStateOpen Then DB715.Close
DB715.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

If RS715.State = adStateOpen Then RS715.Close
RS715.Open "SELECT * FROM [MasterModify] where OldCode = '" & Trim(Text1.Text) & "' OR NewCode= '" & Trim(Text1.Text) & "'", DB715, adOpenKeyset, adLockReadOnly
If RS715.EOF Then
Call displayreport
Else
RS715.MoveFirst
Do Until RS715.EOF
M1AAA = RS715!NewDescription
M2AAA = RS715![Date_Modify]
M3AAA = RS715!NewStock
M4AAA = RS715!NEWMEASURE
M5AAA = RS715!Price
M6AAA = RS715!Update

'Dim RS400 As New ADODB.Recordset
If RS716.State = adStateOpen Then RS716.Close
RS716.Open "SELECT* FROM [HISTORY]", DB715, 3, 3
RS716.AddNew
RS716!Description = M1AAA
RS716!HDate = M2AAA
RS716!Stock = M3AAA
RS716!MEASURE = M4AAA
RS716!Price = M5AAA
RS716!Operation = M6AAA
RS716.Update
RS716.Close
Set RS716 = Nothing
RS715.MoveNext
Loop
RS715.Close
Set RS715 = Nothing
Call displayreport
End If
End Sub
Public Sub displayreport()
Dim DB1112 As New ADODB.CONNECTION
Dim RS1112 As New ADODB.Recordset
Dim RS1113 As New ADODB.Recordset
If DB1112.State = adStateOpen Then DB1112.Close
DB1112.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "HISTORY", acViewPreview
APPACCESS.Visible = True

'DELETE THE TABLE
RS1113.Open "DELETE * FROM HISTORY", DB1112, 3, 3


End Sub
Private Sub Command2_Click()
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim SQL As String
Dim RS12 As New ADODB.Recordset
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table
RS.Open "insert into [RECEIVEDREPORT] SELECT Code, Description, Location, OldStock, OldPrice,Price,Qsupply,Amount, RDate,New_Stock,SuppName,SuppContact,SuppComp,MEASURE FROM [Received] where RDate between #" & Str(DTPicker3.Value) & "# and #" & Str(DTPicker4.Value) & "# ORDER BY ID ", DB, 3, 3

'''coming  OldPrice,
'RS.Open "insert into [SOLDREPORT] SELECT Code, Description, Location, Stock, Price, CustName, TDate,Quantity,Discount,GoodsType,Amount,MEASURE FROM [Sold] where TDate between #" & Str(DTPicker1.Value) & "# and #" & Str(DTPicker2.Value) & "# ORDER BY ID", DB, 3, 3

'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Received", acViewPreview
APPACCESS.Visible = True

'DELETE THE TABLE
RS12.Open "DELETE * FROM RECEIVEDREPORT", DB, 3, 3
End Sub

Private Sub Command3_Click()
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String

DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table
RS.Open "insert into [MASTERMODIFYREPORT] SELECT OldCode, NewCode,OldDescription,NewDescription, OldLocation,NewLocation, OldStock,NewStock, OldPrice,NewPrice,Date_Modify,OLDMEASURE,NEWMEASURE FROM [MasterModify] where Date_Modify between #" & Str(DTPicker5.Value) & "# and #" & Str(DTPicker6.Value) & "# ORDER BY ID", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "MasterModify", acViewPreview
APPACCESS.Visible = True
'DELETE TABLE
RS12.Open "DELETE * FROM MASTERMODIFYREPORT", DB, 3, 3

End Sub

Private Sub Command4_Click()
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table
RS.Open "insert into [MASTERREPORT] SELECT Code ,Description,GoodsType,Location,Stock,Price,DateModify,UNIT FROM [Master] where GoodsType = '" & Combo1.Text & "' ORDER BY DESCRIPTION ASC", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Master", acViewPreview
APPACCESS.Visible = True

'DELETE TABLE
RS12.Open "DELETE * FROM MASTERREPORT", DB, 3, 3
End Sub

Private Sub Command5_Click()
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table
RS.Open "insert into [MASTERREPORT] SELECT Code ,Description,Location,GoodsType,Stock,Price,DateModify,UNIT FROM [Master] ORDER BY GoodsType,Description ASC", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Master", acViewPreview
APPACCESS.Visible = True

'DELETE TABLE
RS12.Open "DELETE * FROM MASTERREPORT", DB, 3, 3
End Sub

Private Sub Command6_Click()
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

Dim REMAIN1
REMAIN1 = 0


'query to filter the table

RS.Open "insert into [MASTERREPORT] SELECT Code ,Description,Location,GoodsType,Stock,Price,DateModify,UNIT FROM [Master] where Stock > " & REMAIN1 & " ORDER BY GoodsType,DESCRIPTION ASC", DB, 3, 3


'RS.Open "insert into [MASTERREPORT] SELECT Code ,Description,GoodsType,Location,Stock,Price,DateModify,UNIT FROM [Master] where Stock < @" & REMAIN & " ORDER BY GoodsType,DESCRIPTION ASC", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Master", acViewPreview
APPACCESS.Visible = True

'DELETE TABLE
RS12.Open "DELETE * FROM MASTERREPORT", DB, 3, 3
End Sub

Private Sub Command7_Click()
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table
RS.Open "insert into [SOLDREPORT] SELECT Code, Description, Location, Stock, Price, CustName, TDate,Quantity,Discount,GoodsType,Amount,MEASURE FROM [Sold] where TDate between #" & Str(DTPicker1.Value) & "# and #" & Str(DTPicker2.Value) & "# ORDER BY ID", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Sold", acViewPreview
APPACCESS.Visible = True
'DELETE TABLE
RS12.Open "DELETE * FROM SOLDREPORT ", DB, 3, 3
End Sub

Private Sub Command8_Click()
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

'query to filter the table
RS.Open "insert into [MASTERREPORT] SELECT Code ,Description,GoodsType,Location,Stock,Price,DateModify,UNIT FROM [Master] where Location = '" & Combo3.Text & "' ORDER BY DESCRIPTION ASC", DB, 3, 3


'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Master", acViewPreview
APPACCESS.Visible = True

'DELETE TABLE
RS12.Open "DELETE * FROM MASTERREPORT", DB, 3, 3
End Sub

Private Sub Command9_Click()
If Text2.Text = "" Then
MsgBox "PLEASE ENTER A VALID NUMBER ..... TO VERIFY GOODS"
Text2.SetFocus
Exit Sub
End If
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS419 As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim SQL As String
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
'Dim REMAIN
'REMAIN = 1
If Option1.Value = True Then
'query to filter the table

RS.Open "insert into [MASTERREPORT] SELECT Code ,Description,Location,GoodsType,Stock,Price,DateModify,UNIT FROM [Master] where Stock >= " & Val(Text2.Text) & " ORDER BY GoodsType,DESCRIPTION ASC", DB, 3, 3


'RS.Open "insert into [MASTERREPORT] SELECT Code ,Description,GoodsType,Location,Stock,Price,DateModify,UNIT FROM [Master] where Stock < @" & REMAIN & " ORDER BY GoodsType,DESCRIPTION ASC", DB, 3, 3
Else

RS419.Open "insert into [MASTERREPORT] SELECT Code ,Description,Location,GoodsType,Stock,Price,DateModify,UNIT FROM [Master] where Stock <= " & Val(Text2.Text) & " ORDER BY GoodsType,DESCRIPTION ASC", DB, 3, 3

End If

'display report
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
APPACCESS.DoCmd.OpenReport "Master", acViewPreview
APPACCESS.Visible = True

'DELETE TABLE
RS12.Open "DELETE * FROM MASTERREPORT", DB, 3, 3

End Sub

Private Sub Form_Load()
Option1.Value = True
Me.Top = 800
Me.Left = 4500
DTPicker1.Value = Format(Now, "DD/MM/YYYY")
DTPicker2.Value = Format(Now, "DD/MM/YYYY")
DTPicker3.Value = Format(Now, "DD/MM/YYYY")
DTPicker4.Value = Format(Now, "DD/MM/YYYY")
DTPicker5.Value = Format(Now, "DD/MM/YYYY")
DTPicker6.Value = Format(Now, "DD/MM/YYYY")
'DTPicker7.Value = Format(Now, "DD/MM/YYYY")
'DTPicker8.Value = Format(Now, "DD/MM/YYYY")

Dim DB2 As New ADODB.CONNECTION
Dim RS4 As New ADODB.Recordset
Dim rs501 As New ADODB.Recordset
Dim rs502 As New ADODB.Recordset
Dim rs503 As New ADODB.Recordset
DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

RS4.Open "SELECT * FROM GoodsType", DB2, adOpenKeyset, adLockReadOnly
RS4.MoveFirst
Do Until RS4.EOF
Combo1.AddItem RS4!Type

Combo7.AddItem RS4!Type
Combo5.AddItem RS4!Type

Combo2.AddItem RS4!Type
RS4.MoveNext
Loop
RS4.Close
Set RS4 = Nothing

rs501.Open "SELECT * FROM Location", DB2, adOpenKeyset, adLockReadOnly
rs501.MoveFirst
Do Until rs501.EOF
'Combo1.AddItem RS4!Type
Combo3.AddItem rs501!Location
rs501.MoveNext
Loop
rs501.Close
Set rs501 = Nothing

'rs502.Open "SELECT * FROM master", DB2, adOpenKeyset, adLockPessimistic
'rs502.MoveFirst
'Do Until rs502.EOF
'Dim q
'rs503.Open "SELECT * FROM master", DB2, adOpenDynamic, adLockOptimistic
'q = Val(rs502!Price) * Val(rs502!Stock)
'rs503.AddNew
'rs503!AllAmount = q
'rs503.Update
'rs502.MoveNext
'Loop
'rs502.Close
'Set rs502 = Nothing

DB2.Close
Set DB2 = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
'Beep
End If
End Sub
