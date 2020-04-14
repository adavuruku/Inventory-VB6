VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMRECEIVE 
   BackColor       =   &H00004000&
   Caption         =   "RECEIVE GOODS FORMS"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16770
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   16770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12840
      TabIndex        =   37
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox txtselprice 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12840
      TabIndex        =   35
      Top             =   4680
      Width           =   3495
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
      Left            =   240
      TabIndex        =   33
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11280
      TabIndex        =   32
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtpriceq 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6120
      TabIndex        =   30
      Top             =   6120
      Width           =   3735
   End
   Begin VB.TextBox txtgoods 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6120
      TabIndex        =   28
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtsearch 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   27
      Top             =   1800
      Width           =   3735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   240
      TabIndex        =   26
      Top             =   2520
      Width           =   3735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5280
      Top             =   6840
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
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "START NEW TRANSACTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   25
      Top             =   7080
      Width           =   3495
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE RECORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   24
      Top             =   6360
      Width           =   3495
   End
   Begin VB.TextBox txtsupcontact 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12840
      TabIndex        =   22
      Text            =   "[NIL]"
      Top             =   4080
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   12840
      TabIndex        =   19
      Top             =   2280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   79626241
      CurrentDate     =   41662
   End
   Begin VB.TextBox txtprice 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12840
      TabIndex        =   13
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtquantity 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12840
      TabIndex        =   12
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txtamount 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtsupcomp 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12840
      TabIndex        =   10
      Text            =   "[NIL]"
      Top             =   2880
      Width           =   3495
   End
   Begin VB.TextBox txtsupname 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12840
      TabIndex        =   9
      Text            =   "[NIL]"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox txtcode 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6120
      TabIndex        =   4
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6120
      TabIndex        =   3
      Top             =   3360
      Width           =   3735
   End
   Begin VB.TextBox txtlocation 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox txtstock 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROFIT EXPECTED (N) :"
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
      Height          =   375
      Left            =   9960
      TabIndex        =   38
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SELING PRICE (N) "
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
      Height          =   375
      Left            =   10200
      TabIndex        =   36
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AVAILABLE ITEMS"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE (N) :"
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
      Height          =   375
      Left            =   4560
      TabIndex        =   31
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GOODS TYPE:"
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
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "RECEIVE DETAILS AND RECORDS"
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
      Height          =   1215
      Left            =   4560
      TabIndex        =   23
      Top             =   480
      Width           =   5895
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER NAME"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   21
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER CONTACT  "
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
      Height          =   375
      Left            =   9720
      TabIndex        =   20
      Top             =   4080
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY "
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
      Height          =   375
      Left            =   11040
      TabIndex        =   18
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT(N) "
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
      Height          =   375
      Left            =   10920
      TabIndex        =   17
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DATE "
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
      Height          =   375
      Left            =   11640
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIVE PRICE (N) "
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
      Height          =   375
      Left            =   10440
      TabIndex        =   15
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER COMPANY"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   14
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      Height          =   1815
      Left            =   11040
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODE :"
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
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION : "
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
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK :"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NAME :"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Height          =   7815
      Left            =   4320
      Top             =   240
      Width           =   12375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   "Label16"
      Height          =   495
      Left            =   360
      TabIndex        =   34
      Top             =   480
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Height          =   7695
      Left            =   120
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "FRMRECEIVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub cmdnewRecord_Click()
'If txtcode.Text = "" Or txtlocation.Text = "" Or txtstock.Text = "" Or txtname.Text = "" _
'Or txtprice.Text = "" Or txtquantity.Text = "" Or txtamount.Text = "" _
'Then
'MsgBox "PLEASE PROVIDE ALL THE REQUIRED DETAILS BEFORE SAVING RECORD"
'Exit Sub
'ElseIf txtstock.Text = txtquantity.Text Then
'Dim DB As New ADODB.CONNECTION
'Dim RS As New ADODB.Recordset
'Dim RS2 As New ADODB.Recordset
'DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
'RS.Open "SELECT * FROM Received", DB, adOpenDynamic, adLockOptimistic
'RS.AddNew
'RS!Code = txtcode
'RS!Description = txtname
'RS!Price = txtprice
'RS!QSupply = txtquantity
'RS!Amount = txtamount
'RS!RDate = DTPicker1.Value
'RS!SuppName = txtsupname
'RS!SuppComp = txtsupcomp
'RS!SuppContact = txtsupcontact
'RS!OldStock = 0
'RS!New_Stock = txtquantity
'RS.Update
'RS.Close
'Set RS = Nothing
'RS2.Open "SELECT * FROM master", DB, adOpenDynamic, adLockOptimistic
'RS2.AddNew
'RS2!Code = txtcode
'RS2!Price = 0
'RS2!Description = txtname
'RS2!Location = txtlocation
'RS2!Stock = txtstock
'RS2!DateModify = DTPicker1.Value
'RS2.Update
'DB.Close
'Set DB = Nothing
'MsgBox "New Record saved successfuly...!!", vbInformation, "Record saved"
'Call clear
'Else
'MsgBox "stock and quantity must be Thesame for registering new Records", vbInformation, "WARNING!!"
'txtquantity.SetFocus
'txtquantity.SelStart = 0
'txtquantity.SelLength = Len(txtquantity.Text)
'Exit Sub
'End If
'End Sub

Private Sub cmdRefresh_Click()
Dim m
m = MsgBox("do you really want to start new Transaction", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If m = vbNo Then
Exit Sub
Else
'cmdnewRecord.Enabled = True
cmdUpdate.Enabled = False
Combo1.Enabled = True
List1.Enabled = True
txtsearch.Locked = False

Call clear
End If
End Sub

Private Sub cmdUpdate_Click()
If txtprice.Text = "" Or txtquantity.Text = "" Or txtamount.Text = "" Then
MsgBox "PLEASE FILL ALL THE REQUIRED DATAS BEFORE UPDATING RECORD", vbCritical, "SOME PARAMETERS LEFT EMPTY"
Exit Sub
Else
Dim nstock
nstock = Val(txtstock.Text) + Val(txtquantity)
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS.Open "SELECT * FROM Received", DB, adOpenDynamic, adLockOptimistic
RS.AddNew
RS!Code = txtcode
RS!Description = txtname
RS!PriceBought = txtprice.Text
RS!Price = txtselprice.Text
RS!QSupply = txtquantity
RS!amount = txtamount
RS!RDate = DTPicker1.Value
RS!SuppName = txtsupname
RS!SuppComp = txtsupcomp
RS!SuppContact = txtsupcontact
RS!OldStock = txtstock
RS!New_Stock = nstock
RS!MEASURE = Combo1.Text
RS!OldPrice = txtpriceq.Text
'RS!OldPrice = txtpriceq.Text
RS!GoodsType = txtgoods.Text
RS!RECEIVED = "RECEIVED"
RS!PROFIT = Val(Text1.Text)
RS.Update
RS.Close
Set RS = Nothing
RS2.Open "SELECT * FROM master WHERE Code = '" & txtcode.Text & "'", DB, adOpenDynamic, adLockOptimistic
If Not RS2.EOF Then
If txtselprice.Text = "" Then
txtselprice.Text = txtpriceq.Text
End If
RS2!Stock = nstock
RS2!DateModify = DTPicker1.Value
RS2!Price = txtselprice.Text
RS2!PROFIT = Val(Text1.Text)
RS2.Update
Else
Exit Sub
End If
DB.Close
Set DB = Nothing
MsgBox "Record SAVED Succesfully and the stock fully Updated!!!", vbInformation, "SUCCESFUL UPDATE"
cmdUpdate.Enabled = False
'cmdnewRecord.Enabled = True
Combo1.Enabled = True
List1.Enabled = True
txtsearch.Locked = False
Call clear
End If
End Sub
 
Private Sub DataList1_Click()

End Sub




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

Private Sub Command1_Click()
Dim J
J = MsgBox("Do you really want to CLOSE THIS RECEIVE FORM..PLEASE VERIFY you may lost the UNSAVED process..please confirm !!", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If J = vbNo Then
MsgBox "for closing Denied ...you can continue with your process !!!!!"
Exit Sub
Else
Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Top = 1900
Me.Left = 3100
txtcode.Locked = True
txtname.Locked = True
txtlocation.Locked = True
txtstock.Locked = True
txtgoods.Locked = True
txtpriceq.Locked = True
DTPicker1.Value = Format(Now, "DD/MM/YYYY")
cmdUpdate.Enabled = False
Dim DB2 As New ADODB.CONNECTION
Dim RS3 As New ADODB.Recordset
Dim RS322 As New ADODB.Recordset
Dim RS326 As New ADODB.Recordset
If DB2.State = adStateOpen Then DB2.Close
DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS3.Open "SELECT * FROM master", DB2, adOpenKeyset, adLockReadOnly
RS3.MoveFirst
Do Until RS3.EOF
List1.AddItem RS3!Description
'& " = " & RS3!GoodsType
RS3.MoveNext
Loop
RS3.Close
Set RS3 = Nothing

If RS322.State = adStateOpen Then RS322.Close
RS322.Open "SELECT * FROM UNIT", DB2, adOpenKeyset, adLockReadOnly
RS322.MoveFirst
Do Until RS322.EOF
Combo1.AddItem RS322!measures
RS322.MoveNext
Loop
RS322.Close
Set RS322 = Nothing

'If RS326.State = adStateOpen Then RS326.Close
'RS326.Open "SELECT * FROM UNIT", DB2, adOpenKeyset, adLockReadOnly
'RS326.MoveFirst
'Combo1.Text = RS326!measures
'RS326.Close
'Set RS326 = Nothing

DB2.Close
Set DB2 = Nothing
End Sub

Private Sub List1_Click()
On Error Resume Next
txtcode.Locked = True
txtname.Locked = True
txtlocation.Locked = True
txtstock.Locked = True
txtgoods.Locked = True
txtpriceq.Locked = True
'txtsprice.Locked = True
cmdUpdate.Enabled = True
'cmdUpdate.Enabled = True
Dim DB As New ADODB.CONNECTION
Dim RS As New ADODB.Recordset
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS.Open "SELECT * FROM master where Description = '" & List1.Text & "'", DB, adOpenDynamic, adLockOptimistic
txtcode = RS!Code
txtname = RS!Description
txtlocation = RS!Location
txtstock = RS!Stock
txtpriceq.Text = RS!Price
txtgoods.Text = RS!GoodsType
txtselprice.Text = RS!Price
RS.Close
Set RS = Nothing
DB.Close
Set DB = Nothing
Combo1.Enabled = False
'List1.Enabled = False
txtsearch.Locked = True
End Sub

Private Sub txtprice_Change()
txtamount = Val(txtprice.Text) * Val(txtquantity.Text)
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
'Beep
End If
End Sub

Private Sub txtquantity_Change()
txtamount = Val(txtprice.Text) * Val(txtquantity.Text)
End Sub

Private Sub txtsprice_Change()

End Sub

Private Sub txtquantity_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
'Beep
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

Private Sub txtselprice_Change()
Text1.Text = (Val(txtselprice.Text) * Val(txtquantity.Text)) - Val(txtamount)
End Sub

Private Sub txtselprice_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
'Beep
End If
End Sub

Private Sub txtstock_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
End If

End Sub

