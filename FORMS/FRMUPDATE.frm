VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMUPDATE 
   BackColor       =   &H00004000&
   Caption         =   "UPDATE RECORD IN STOCKS"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "DELLETE / REMOVE GOODS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   7
      Top             =   7800
      Width           =   6255
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
      Left            =   7440
      TabIndex        =   3
      Top             =   3240
      Width           =   3855
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
      Left            =   360
      TabIndex        =   22
      Top             =   1200
      Width           =   3735
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
      Left            =   7440
      TabIndex        =   2
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox txtdate 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5160
      Width           =   3855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   360
      TabIndex        =   19
      Top             =   2640
      Width           =   3735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
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
   Begin VB.TextBox txtsearch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE RECORD"
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
      Left            =   8520
      TabIndex        =   9
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START NEW UPDATE"
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
      Left            =   4920
      TabIndex        =   8
      Top             =   7080
      Width           =   2895
   End
   Begin VB.TextBox txtstock 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7440
      TabIndex        =   5
      Top             =   4560
      Width           =   3855
   End
   Begin VB.TextBox txtprice 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7440
      TabIndex        =   4
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox txtcode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   1920
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   7440
      TabIndex        =   10
      Top             =   5760
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   78643201
      CurrentDate     =   41662
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MEASUREMENT (UNIT): "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   3360
      Width           =   3375
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "UPDATE DETAILS WINDOW"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      TabIndex        =   21
      Top             =   360
      Width           =   6615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LAST DATE UPDATED "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      Height          =   1815
      Left            =   4440
      Top             =   6840
      Width           =   7215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOW UPDATE DATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL GOOD IN STOCK "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GOODS PRICE (N) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GOODS CODE "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION OF GOODS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   2760
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GOODS NAME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Height          =   6615
      Left            =   4320
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT GOODS TO UPDATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Height          =   8415
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "FRMUPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
List1.clear
Dim DB117 As New ADODB.CONNECTION
Dim RS117 As New ADODB.Recordset
DB117.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS117.Open "SELECT * FROM master where unit = '" & Combo1.Text & "'", DB117, 3, 3
Do Until RS117.EOF
List1.AddItem RS117!Description
RS117.MoveNext
Loop
RS117.Close
Set RS117 = Nothing
DB117.Close
Set DB117 = Nothing
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
m = MsgBox("do you really want to start new Transaction", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If m = vbNo Then
Exit Sub
Else
Call uclear
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim QY
QY = MsgBox("Do you really want to UPDATE THE SELECTED GOODS ...Please confirm", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If QY = vbNo Then
Exit Sub
Else
Dim DB As New ADODB.CONNECTION
Dim RS3 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS2.Open "SELECT * FROM MasterModify", DB, adOpenDynamic, adLockOptimistic
RS2.AddNew
RS2!NewCode = txtcode
RS2!NewDescription = txtname
RS2!NewPrice = txtprice
RS2!NewLocation = txtlocation
RS2!NewStock = txtstock
RS2!NEWMEASURE = Combo3.Text
RS2!OldPrice = old_price
RS2!OldStock = old_stock
RS2!OldCode = old_code
RS2!OldLocation = old_location
RS2!OldDescription = old_name
RS2!OLDMEASURE = OLDMEASURESD
'RS2!OldRecdate = old_date
RS2!Date_Modify = DTPicker1.Value
RS2!Update = "Update"
RS2!GoodsType = fff
RS2.Update
RS2.Close
Set RS2 = Nothing
RS3.Open "SELECT * FROM master WHERE Code = '" & txtcode.Text & "'", DB, adOpenDynamic, adLockOptimistic
If Not RS3.EOF Then
RS3!Code = txtcode
RS3!Description = txtname
RS3!Price = txtprice
RS3!Location = Combo2.Text
RS3!Stock = txtstock
RS3!unit = Combo3.Text
'RS3!DateModify = DTPicker1.Value
RS3.Update
Else
Exit Sub
End If
DB.Close
Set DB = Nothing
MsgBox "The Record Was Succesfully UPDATED", vbOKOnly + vbInformation, "RECORD UPDATE SUCCESFULL"

txtprice.Enabled = False
txtcode.Enabled = False
txtstock.Enabled = False
txtname.Enabled = False
txtdate.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Command2.Enabled = False
Command3.Enabled = False

Call uclear
End If
End Sub

Private Sub Command3_Click()
Dim Qpo
Qpo = MsgBox("Do you really want to DELETE / REMOVE THE SELECTED GOODS FROM GOODS AVAILABLE IN THE STORE...Please confirm", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If Qpo = vbNo Then
Exit Sub
Else
Dim DB201 As New ADODB.CONNECTION
Dim RS377 As New ADODB.Recordset
Dim RS301 As New ADODB.Recordset
If DB201.State = adStateOpen Then DB201.Close
DB201.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS301.Open "DELETE * FROM master where Code ='" & Trim(txtcode.Text) & "'", DB201, 3, 3

RS377.Open "SELECT * FROM master", DB201, adOpenKeyset, adLockReadOnly
List1.clear
If RS377.EOF Or RS377.BOF Then
List1.clear
Exit Sub
Else
RS377.MoveFirst
Do Until RS377.EOF
List1.AddItem RS377!Description
RS377.MoveNext
Loop
RS377.Close
Set RS377 = Nothing

DB201.Close
Set DB201 = Nothing

MsgBox "RECORD HAS BEEN REMOVED / DELETED SUCCESFULLY"

txtprice.Enabled = False
txtcode.Enabled = False
txtstock.Enabled = False
txtname.Enabled = False
txtdate.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Command2.Enabled = False
Command3.Enabled = False

Call uclear

End If
End If
End Sub

Private Sub Form_Load()
Me.Top = 1400
Me.Left = 5000
txtprice.Enabled = False
txtcode.Enabled = False
txtstock.Enabled = False
txtname.Enabled = False
txtdate.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
DTPicker1.Value = Format(Now, "DD/MM/YYYY")
Command2.Enabled = False
Command3.Enabled = False
Dim DB2 As New ADODB.CONNECTION
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim RS322 As New ADODB.Recordset
Dim RS326 As New ADODB.Recordset
If DB2.State = adStateOpen Then DB2.Close
DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS3.Open "SELECT * FROM master", DB2, adOpenKeyset, adLockReadOnly
RS3.MoveFirst
Do Until RS3.EOF
List1.AddItem RS3!Description
RS3.MoveNext
Loop
RS3.Close
Set RS3 = Nothing

RS4.Open "SELECT * FROM Location", DB2, adOpenKeyset, adLockReadOnly
RS4.MoveFirst
Do Until RS4.EOF
Combo2.AddItem RS4!Location
RS4.MoveNext
Loop
RS4.Close
Set RS4 = Nothing

If RS322.State = adStateOpen Then RS322.Close
RS322.Open "SELECT * FROM UNIT", DB2, adOpenKeyset, adLockReadOnly
RS322.MoveFirst
Do Until RS322.EOF
Combo1.AddItem RS322!measures
Combo3.AddItem RS322!measures
RS322.MoveNext
Loop
RS322.Close
Set RS322 = Nothing

If RS326.State = adStateOpen Then RS326.Close
RS326.Open "SELECT * FROM UNIT", DB2, adOpenKeyset, adLockReadOnly
RS326.MoveFirst
Combo1.Text = RS326!measures
RS326.Close
Set RS326 = Nothing

DB2.Close
Set DB2 = Nothing
End Sub

Private Sub List1_Click()
txtcode.Text = ""
txtname.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
txtprice.Text = ""
txtstock.Text = ""
txtdate.Text = ""
On Error Resume Next
'txtsearch.Text = List1.Text
txtprice.Enabled = True
txtcode.Enabled = False
txtstock.Enabled = True
txtname.Enabled = True
txtdate.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Command3.Enabled = True

Command2.Enabled = True
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS.Open "SELECT * FROM master where Description = '" & List1.Text & "'", DB, adOpenDynamic, adLockOptimistic
txtcode = RS!Code
txtname = RS!Description
Combo2.Text = RS!Location
txtstock = RS!Stock
txtprice = RS!Price
txtdate = RS!DateModify
Combo3.Text = RS!unit
fff = RS!GoodsType
RS.Close
Set RS = Nothing
DB.Close
Set DB = Nothing
old_price = txtprice
old_stock = txtstock
old_code = txtcode
old_location = txtlocation
old_name = txtname
old_date = txtdate
OLDMEASURESD = Combo3.Text
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtsearch_Change()
List1.clear
Dim DB11 As New ADODB.CONNECTION
Dim RS11 As New ADODB.Recordset
DB11.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS11.Open "SELECT * FROM master where Description LIKE '%" & Trim(txtsearch.Text) & "%' AND unit = '" & Combo1.Text & "'", DB11, 3, 3
Do Until RS11.EOF
List1.AddItem RS11!Description
RS11.MoveNext
Loop
RS11.Close
Set RS11 = Nothing
DB11.Close
Set DB11 = Nothing

End Sub

Private Sub txtstock_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
