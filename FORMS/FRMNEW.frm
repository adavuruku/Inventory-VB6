VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMNEW 
   BackColor       =   &H00004000&
   Caption         =   "NEW GOODS "
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      TabIndex        =   2
      Top             =   3240
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   10
      Top             =   6000
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   6480
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
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   9
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox txtprice 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      TabIndex        =   4
      Top             =   4920
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtstock 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   4080
      Width           =   4455
   End
   Begin VB.TextBox txtcode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MEASUREMENT UNIT:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "NEW GOODS REGISTRATION RECORDS"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      Height          =   1095
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   8775
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE (N) :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GOODS TYPE:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NAME :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   4920
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODE :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      Height          =   6975
      Left            =   120
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "FRMNEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
GOOD = ""
GOOD2 = ""
GOOD3 = ""
SGOOD = ""
GOOD4 = ""
Dim DB2 As New ADODB.CONNECTION
Dim RS4 As New ADODB.Recordset
DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
SGOOD = Trim(Combo1.Text)
SQL = "SELECT * FROM GoodsType where Type='" & SGOOD & "'"
RS4.Open SQL, DB2, 3, 3

GOOD = RS4!GOODSCODE
GOOD4 = RS4!REP
GOOD2 = Val(GOOD) + 1
GOOD3 = GOOD4 & GOOD2
txtcode.Text = GOOD3
DB2.Close
Set DB2 = Nothing
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
Dim m
m = MsgBox("Do you really want to SAVE THE NEW Transaction", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If m = vbNo Then
Combo1.SetFocus
Exit Sub
Else
If txtcode.Text = "" Or Combo1.Text = "" Or txtname.Text = "" _
Or Combo2.Text = "" Or Combo3.Text = "" Then
MsgBox "PLEASE PROVIDE ALL THE REQUIRED DETAILS BEFORE SAVING NEW RECORDs", vbCritical, "ERROR MESSAGE"
Exit Sub
Else
Dim DB As New ADODB.CONNECTION
Dim RS4 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS2.Open "SELECT * FROM master", DB, adOpenDynamic, adLockOptimistic
RS2.AddNew
RS2!Code = txtcode
RS2!Price = "0"
RS2!Description = txtname
RS2!Location = Combo2.Text
RS2!Stock = "0"
RS2!unit = Combo3.Text
RS2!DateReg = Format(Now, "DD/MM/YYYY")
RS2!New = "New Reg"
RS2!GoodsType = Combo1.Text
RS2.Update
RS2.Close
Set RS2 = Nothing

'..................UPDATING THE CODE FOR NEXT REGISTRATION................
RS4.Open "SELECT * FROM GoodsType WHERE Type = '" & Combo1.Text & "'", DB, adOpenDynamic, adLockOptimistic
If Not RS4.EOF Then
RS4!GOODSCODE = GOOD2
RS4.Update
Else
Exit Sub
End If

MsgBox "New Record Created and saved successfuly...!!", vbInformation, "Record saved"
txtcode.Text = ""
txtprice.Text = ""
txtname.Text = ""
Combo2.Text = ""
Combo1.Text = ""
txtstock.Text = ""
Combo3.Text = ""

End If
End If
DB.Close
Set DB = Nothing
End Sub

'!..........................

'RS4.Open "SELECT * FROM GoodsType WHERE Type = '" & SGOOD & "'", DB2, 3, 3
'If Not RS4.EOF Then
'RS4!GOODSCODE = GOOD2
'RS4.Update
'Else
'End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!



Private Sub Command2_Click()
Dim m
m = MsgBox("do you really want to start new Transaction", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If m = vbNo Then
Combo1.SetFocus
Exit Sub
Else
txtcode.Text = ""
txtprice.Text = ""
txtname.Text = ""
Combo2.Text = ""
Combo1.Text = ""
txtstock.Text = ""
End If
End Sub

Private Sub Command3_Click()
Dim J
J = MsgBox("Do you really want to CLOSE THIS NEW GOODS REGISTRATION FORM..PLEASE VERIFY you may lost the UNSAVED process..please confirm !!", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If J = vbNo Then
MsgBox "Form closing Has been Denied by you ...you can continue with your process !!!!!"
Exit Sub
Else
Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Top = 2000
Me.Left = 5500
Dim DB2 As New ADODB.CONNECTION
Dim RS4 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS322 As New ADODB.Recordset
If DB2.State = adStateOpen Then DB2.Close
DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS3.Open "SELECT * FROM Location", DB2, adOpenKeyset, adLockReadOnly
RS3.MoveFirst
Do Until RS3.EOF
Combo2.AddItem RS3!Location
RS3.MoveNext
Loop
RS3.Close
Set RS3 = Nothing

RS4.Open "SELECT * FROM GoodsType", DB2, adOpenKeyset, adLockReadOnly
RS4.MoveFirst
Do Until RS4.EOF
Combo1.AddItem RS4!Type
RS4.MoveNext
Loop
RS4.Close
Set RS4 = Nothing

If RS322.State = adStateOpen Then RS322.Close
RS322.Open "SELECT * FROM UNIT", DB2, adOpenKeyset, adLockReadOnly
RS322.MoveFirst
Do Until RS322.EOF
Combo3.AddItem RS322!measures
RS322.MoveNext
Loop
RS322.Close
Set RS322 = Nothing

DB2.Close
Set DB2 = Nothing
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

Private Sub txtstock_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub
