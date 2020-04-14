VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   Caption         =   "Form2"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14145
   LinkTopic       =   "Form2"
   ScaleHeight     =   9390
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE CHANGES"
      Height          =   615
      Left            =   8640
      TabIndex        =   1
      Top             =   7560
      Width           =   5295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5280
      Top             =   360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16776960
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim adoconn As New ADODB.Connection
Dim RS3 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim str1 As String


Private Sub Command1_Click()
rs1.Update
End Sub

Private Sub Form_Load()

If adoconn.State = adStateOpen Then adoconn.Close
adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
If rs1.State = adStateOpen Then rs1.Close

adoconn.CursorLocation = adUseClient
    ' bring to paste
str1 = "SELECT Stand,Elongation,constant,constanteven,Diameter_roll,SValue,Transmission_Ratio,D_Value,Mn,C,Q FROM Maths2"
    If rs1.State = adStateOpen Then rs1.Close
'    rs1.Open str1, adoconn, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = rs1
End Sub
