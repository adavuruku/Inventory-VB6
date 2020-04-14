VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMSOLD 
   BackColor       =   &H00004000&
   Caption         =   "SALES FORM"
   ClientHeight    =   8055
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4680
      Left            =   18720
      TabIndex        =   37
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4680
      Left            =   17640
      TabIndex        =   36
      Top             =   2400
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   35
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4680
      Left            =   15360
      TabIndex        =   33
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtsearch2 
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
      Left            =   15360
      TabIndex        =   32
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   10
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE SALES"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   9
      Top             =   5280
      Width           =   2415
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
      Left            =   240
      TabIndex        =   31
      Top             =   1800
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   240
      TabIndex        =   30
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "START NEW TRANSACTION"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   7
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4560
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT RECEIPT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   8
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT TRANSACTION"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11280
      TabIndex        =   6
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox txtlocation 
      BorderStyle     =   0  'None
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
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   2880
      Width           =   3495
   End
   Begin VB.TextBox txtname 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtcode 
      BorderStyle     =   0  'None
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox txtdiscount 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   4
      Text            =   "0"
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox txtquantity 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   3
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox txtcustname 
      BorderStyle     =   0  'None
      DataField       =   "CustName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   1
      Text            =   "[NIL]"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox txtstock 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3600
      Width           =   3495
   End
   Begin VB.TextBox txtprice 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox txtamount 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   5
      Text            =   "0"
      Top             =   4200
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Bindings        =   "FRMRECEIVE.frx":0000
      Height          =   495
      Left            =   10920
      TabIndex        =   2
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
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
      Format          =   40239105
      CurrentDate     =   41662
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   18840
      TabIndex        =   40
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   17520
      TabIndex        =   39
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF GOODS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   15360
      TabIndex        =   38
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GOODS BOUGHT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15960
      TabIndex        =   34
      Top             =   600
      Width           =   3375
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Height          =   7455
      Left            =   15240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4935
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      BorderWidth     =   10
      Height          =   2295
      Left            =   10920
      Top             =   4920
      Width           =   3855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   3720
      TabIndex        =   29
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GRAND TOTAL PRIZE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3720
      TabIndex        =   28
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      BorderWidth     =   10
      Height          =   2415
      Left            =   5640
      Top             =   4920
      Width           =   9255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "THANKS FOR PATRONIZING US....WE WISH TO SEE YOU AGAIN....WE REALLY APPRECIATE...!!"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   27
      Top             =   7680
      Width           =   14055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      BorderWidth     =   10
      Height          =   2535
      Left            =   3600
      Top             =   4920
      Width           =   11415
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
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
      Height          =   495
      Left            =   4080
      TabIndex        =   26
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION "
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
      Left            =   3840
      TabIndex        =   25
      Top             =   2880
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODE "
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
      TabIndex        =   24
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY :"
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
      Left            =   9360
      TabIndex        =   20
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE (N) "
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
      Left            =   3720
      TabIndex        =   19
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DATE :"
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
      Left            =   9840
      TabIndex        =   18
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT(N) :"
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
      Left            =   9000
      TabIndex        =   17
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK "
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
      Left            =   3720
      TabIndex        =   16
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "BUYER NAME :"
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
      Left            =   9000
      TabIndex        =   15
      Top             =   1560
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DISCOUNT(N) :"
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
      Left            =   8880
      TabIndex        =   14
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "SALES DETAILS AND RECORDS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   13
      Top             =   360
      Width           =   9135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AVAILABLE ITEMS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      BorderWidth     =   9
      Height          =   7455
      Left            =   3600
      Top             =   120
      Width           =   11535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Height          =   7455
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FRMSOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
    Check2.Value = 0
End Sub
Private Sub Check2_Click()
    Check1.Value = 0
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
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Command1_Click()
    Dim gtotal, check, CheckO2
    check = Val(txtstock) - Val(txtquantity)
    CheckO2 = Val(txtamount) - Val(txtdiscount)
    If check < 0 Then
        MsgBox "the Quantity is more than the stock available ..please verify", vbInformation, "Wrong entyr detected"
        txtquantity.SetFocus
        txtquantity.SelStart = 0
        txtquantity.SelLength = Len(txtquantity.Text)
    Exit Sub
    ElseIf CheckO2 < 0 Then
        MsgBox "the Discount price is more than the total amount of the goods customer want or they are equal..please verify", vbInformation, "Wrong entyr detected"
        txtdiscount.SetFocus
        txtdiscount.SelStart = 0
        txtdiscount.SelLength = Len(txtdiscount.Text)
    Exit Sub
    ElseIf txtcustname.Text = "" Then
        MsgBox "please..make sure the customer name is entered..please verify", vbInformation, "Wrong entyr detected"
        txtcustname.SetFocus
        txtcustname.BackColor = vbRed
    Exit Sub
    ElseIf txtquantity.Text = "" Or Val(txtquantity.Text) <= 0 Then
        MsgBox "please..make sure the quantity of goods to buy are entered..please verify... and the value is more than zero(0)", vbInformation, "Wrong entyr detected"
        txtquantity.SetFocus
        txtquantity.BackColor = vbRed
    Exit Sub
    End If

'Command2.Enabled = True


'If Check2.Value = 1 Then

    gtotal = txtamount.Text
    total = Val(total) + Val(gtotal)
    Label15.Caption = "N" & " " & total
    Dim DB As New ADODB.CONNECTION
    Dim RS As New ADODB.Recordset
    Dim RS377 As New ADODB.Recordset
    Dim RS2 As New ADODB.Recordset
    If DB.State = adStateOpen Then DB.Close
        DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
        RS.Open "SELECT * FROM TEMPSOLD", DB, adOpenDynamic, adLockOptimistic
    If txtdiscount.Text = "" Then
        txtdiscount.Text = 0
    End If
        RS.AddNew
        RS!Code = txtcode
        RS!Description = txtname
        RS!Price = txtprice
        RS!Quantity = txtquantity
        RS!amount = txtamount
        RS!TDate = DTPicker1.Value
        RS!Location = txtlocation
        RS!discount = txtdiscount
        RS!custname = txtcustname
        RS!Stock = txtstock
        RS!GoodsType = TYPESS
        RS!MEASURE = Combo1.Text
        RS!sold = "sold"
        'RS!ID = sclick
    RS.Update
    RS.Close
    Set RS = Nothing
    
    ''''CLEAR THE LIST '''''''
    List2.clear
    List3.clear
    List4.clear
    
    ''''POPULATE THE LIST WITH THE NEW SALE'''''''
    
    RS377.Open "SELECT * FROM TEMPSOLD", DB, adOpenKeyset, adLockReadOnly
    RS377.MoveFirst
    Do Until RS377.EOF
        List2.AddItem RS377!Description
        List3.AddItem RS377!Quantity
        List4.AddItem RS377!amount
    RS377.MoveNext
    Loop
    RS377.Close
    Set RS377 = Nothing
    
    DB.Close
    Set DB = Nothing
        MsgBox "Transaction was succesfully completed ..system ready for new transaction..thanks for your patronage", vbInformation, "APPRECIATION MESSAGE"
        List2.Enabled = True
        txtsearch2.Locked = False
        List3.Enabled = True
        List4.Enabled = True
    Call sclear
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    Dim DB11 As New ADODB.CONNECTION
    Dim RS11 As New ADODB.Recordset
    Dim RS12 As New ADODB.Recordset
    Dim DB2 As New ADODB.CONNECTION
    Dim RS3 As New ADODB.Recordset
    Dim RS4 As New ADODB.Recordset
    
    Dim CCODE, QBOUGHT, fstock, J
    J = MsgBox("Do you really want to PRINT RECEIPT For these new Transaction...Please confirm from the customer!!", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
    If J = vbNo Then
        Exit Sub
    Else
        DCLICK = "0"
        sclick = "0"
    If DB2.State = adStateOpen Then DB2.Close
        DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
        RS3.Open "SELECT * FROM TEMPSOLD", DB2, adOpenKeyset, adLockReadOnly
    RS3.MoveFirst
    Do Until RS3.EOF
        CCODE = RS3!Code
        RS4.Open "SELECT * FROM master where Code='" & CCODE & "'", DB2, adOpenDynamic, adLockOptimistic
        If Not RS4.EOF Then
            fstock = RS4!Stock
            RS4!Stock = Val(fstock) - Val(RS3!Quantity)
            RS4!DateModify = Format(Now, "DD/MM/YYYY")
            RS4.Update
        RS4.Close
        Set RS4 = Nothing
    RS3.MoveNext
    Else
    End If
    Loop
    RS3.Close
    Set RS3 = Nothing
    DB2.Close
    Set DB2 = Nothing
    
    'display report DISABLE THE RECEIPT SINCE NO RECEIEPT AVAILABLE
    'WILL BE UNLOCKED WHEN THE PRINTER FOR RECEIPT IS READY
    'REMEMBER
    
    'Dim ACCESSAPP As Access.Application
    'Set APPACCESS = New Access.Application
    'Set APPACCESS = CreateObject("ACCESS.APPLICATION")
    'APPACCESS.OpenCurrentDatabase ("C:\Market\DATABASE\market.mdb")
    'APPACCESS.DoCmd.OpenReport "TEMPSOLD", acViewPreview
    'APPACCESS.Visible = True
    
    'INSERT IN TO SOLD FROM TEMPSOLD
        If DB11.State = adStateOpen Then DB11.Close
        DB11.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
        RS11.Open "insert into SOLD SELECT [Code], [Description], [Location], [Stock], [Price], [CustName], [TDate],[Quantity],[Discount],[Amount],[GoodsType],[MEASURE],[sold] FROM TEMPSOLD", DB11, 3, 3
    
    'DELETE ALL THE RECORDS IN TEMPSOLD
        RS12.Open "DELETE * FROM TEMPSOLD ", DB11, 3, 3
        DB11.Close
        Set DB11 = Nothing
        Label15.Caption = ""
        total = ""
        List2.clear
        List3.clear
        List4.clear
        MsgBox ("TRANSACTION SAVED SUCCESFULLY ..SYSTEM READY FOR NEW TRANSACTIONS")
    End If
    'Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Dim m
m = MsgBox("Do you Really Want to Start a New Transaction", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If m = vbNo Then
    Exit Sub
Else
    Dim DB11 As New ADODB.CONNECTION
    Dim RS12 As New ADODB.Recordset
    DB11.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
    RS12.Open "DELETE * FROM TEMPSOLD ", DB11, 3, 3
        List2.clear
        List3.clear
        List4.clear
        txtsearch.SetFocus
    DB11.Close
    Set DB11 = Nothing
    Label15.Caption = ""
    total = "0"
    Call sclear
    MsgBox "system ready for new transaction !!!", vbInformation, "succes message"
End If
End Sub


Private Sub Command4_Click()
    total = Val(total) - Val(nuh)
    gtotal = txtamount.Text
    total = Val(total) + Val(gtotal)
    Label15.Caption = "N" & " " & total
    'sclick = Val(sclick) + 1
    Dim DB400 As New ADODB.CONNECTION
    Dim RS400 As New ADODB.Recordset
    Dim RS378 As New ADODB.Recordset
    If DB400.State = adStateOpen Then DB400.Close
    DB400.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
    If RS400.State = adStateOpen Then RS.Close
        RS400.Open "SELECT * FROM TEMPSOLD where Code ='" & Trim(txtcode.Text) & "'", DB400, adOpenDynamic, adLockOptimistic
        'RS.AddNew
        RS400!Code = txtcode
        RS400!Description = txtname
        RS400!Price = txtprice
        RS400!Quantity = txtquantity
        RS400!amount = txtamount
        RS400!TDate = DTPicker1.Value
        RS400!Location = txtlocation
        RS400!discount = txtdiscount
        RS400!custname = txtcustname
        RS400!Stock = txtstock
        RS400!GoodsType = TYPESS
        RS400!MEASURE = Combo1.Text
        RS400.Update
    RS400.Close
    Set RS400 = Nothing
    MsgBox "Transaction was succesfully Updated ..system ready for new transaction..thanks for your patronage", vbInformation, "APPRECIATION MESSAGE"
    Command4.Enabled = False
    Command5.Enabled = False
    Command1.Enabled = True
    
    ''''CLEAR THE LIST '''''''
    List2.clear
    List3.clear
    List4.clear
    
    ''''POPULATE THE LIST WITH THE NEW SALE'''''''
    If RS378.State = adStateOpen Then RS378.Close
    RS378.Open "SELECT * FROM TEMPSOLD", DB400, adOpenKeyset, adLockReadOnly
    RS378.MoveFirst
    Do Until RS378.EOF
        List2.AddItem RS378!Description
        List3.AddItem RS378!Quantity
        List4.AddItem RS378!amount
    RS378.MoveNext
    Loop
    RS378.Close
    Set RS378 = Nothing
    
    
    
    List1.Enabled = True
    Combo1.Locked = False
    txtsearch.Locked = False
    Command1.Enabled = False
    
    nuh = "0"
    
    DB400.Close
    Set DB400 = Nothing
    Call sclear
End Sub

Private Sub Command5_Click()
Dim Qpn
Qpn = MsgBox("Do you really want to DELETE / REMOVE THE SELECTED GOODS FROM GOODS AVAILABLE IN THE STORE...Please confirm", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If Qpn = vbNo Then
    Exit Sub
Else
    Dim DB201 As New ADODB.CONNECTION
    Dim RS301 As New ADODB.Recordset
    Dim RS377 As New ADODB.Recordset
    If DB201.State = adStateOpen Then DB201.Close
    DB201.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
    
    RS301.Open "DELETE * FROM TEMPSOLD where Code ='" & Trim(txtcode.Text) & "'", DB201, 3, 3
    
    'Label15.Caption = ""
    total = Val(total) - Val(txtamount)
    Label15.Caption = "N" & " " & total
    Command4.Enabled = False
    Command5.Enabled = False
    Command1.Enabled = True
    List1.Enabled = True
    Combo1.Locked = False
    txtsearch.Locked = False


    'If DB201.State = adStateOpen Then DB201.Close
    RS377.Open "SELECT * FROM TEMPSOLD", DB201, adOpenKeyset, adLockReadOnly
    List2.clear
    List3.clear
    List4.clear
    If RS377.EOF Or RS377.BOF Then
        List2.clear
        Exit Sub
    Else
        RS377.MoveFirst
        Do Until RS377.EOF
            List2.AddItem RS377!Description
            List3.AddItem RS377!Quantity
            List4.AddItem RS377!amount
            RS377.MoveNext
        Loop
        RS377.Close
        Set RS377 = Nothing
        
        DB201.Close
        Set DB201 = Nothing
        Command1.Enabled = False
        Call sclear
    End If
End If
End Sub

Private Sub Command6_Click()

End Sub
Private Sub Form_Load()
    Command4.Enabled = False
    Command5.Enabled = False
    Me.Top = 1300
    Me.Left = 60
    'Check2.Value = 1
    txtcustname.Locked = True
    txtquantity.Locked = True
    txtdiscount.Locked = True
    txtamount.Locked = True
    List2.Enabled = False
    List3.Enabled = False
    List4.Enabled = False
    txtsearch2.Locked = True
    'Command2.Enabled = False
    DTPicker1.Value = Format(Now, "DD/MM/YYYY")
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
    
    If RS326.State = adStateOpen Then RS326.Close
    RS326.Open "SELECT * FROM UNIT", DB2, adOpenKeyset, adLockReadOnly
    RS326.MoveFirst
    Combo1.Text = RS326!measures
    RS326.Close
    Set RS326 = Nothing
    
    DB2.Close
    Set DB2 = Nothing
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
   Exit Sub
Else
   KeyAscii = 0
   Beep
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
    Exit Sub
Else
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub g_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    total = "0"
End Sub

Private Sub List1_Click()
    On Error Resume Next
    txtcustname.Locked = False
    txtquantity.Locked = False
    txtdiscount.Locked = False
    txtamount.Locked = False
    Dim DB As New ADODB.CONNECTION
    Dim RS As New ADODB.Recordset
    DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
    RS.Open "SELECT * FROM master where Description = '" & List1.Text & "'", DB, adOpenDynamic, adLockOptimistic
    If (RS!Price) = "" Then
        MsgBox "PLEASE THE PRICE FOR THESE GOODS IS NOT SET ...TRY TO UPDATE THIS GOODS PRICE BEFORRE SALES", vbInformation, "WRONG TRANSACTION"
        Exit Sub
    Else
        txtcode = RS!Code
        txtname = RS!Description
        txtlocation = RS!Location
        txtstock = RS!Stock
        txtprice = Val(RS!Price)
        TYPESS = RS!GoodsType
        RS.Close
        Set RS = Nothing
        DB.Close
        Set DB = Nothing
        txtamount = "0"
        'txtprice = "0"
        txtquantity = "0"
        txtdiscount = "0"
        Command4.Enabled = False
        Command5.Enabled = False
        Command1.Enabled = True
    End If
End Sub

Private Sub List2_Click()
Command4.Enabled = True
Command5.Enabled = True
Command1.Enabled = False
txtcustname.Locked = False
txtquantity.Locked = False
txtdiscount.Locked = False
txtamount.Locked = False
Dim DB120 As New ADODB.CONNECTION
Dim RS120 As New ADODB.Recordset
Dim RS121 As New ADODB.Recordset
Dim RS122 As New ADODB.Recordset
If DB120.State = adStateOpen Then DB120.Close
DB120.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"

If RS121.State = adStateOpen Then RS121.Close
    RS121.Open "SELECT * FROM TEMPSOLD WHERE Description =  '" & List2.Text & "'", DB120, adOpenDynamic, adLockOptimistic
    If RS121.EOF Then
    Exit Sub
Else
    txtcode = RS121!Code
    txtname = RS121!Description
    txtlocation = RS121!Location
    txtstock = RS121!Stock
    txtprice = Val(RS121!Price)
    txtcustname = RS121!custname
    txtquantity = Val(RS121!Quantity)
    txtdiscount = Val(RS121!discount)
    txtamount = Val(RS121!amount)
    Combo1.Text = RS121!MEASURE
    'asign amount to a public var
    nuh = Val(RS121!amount)
    
    RS121.Close
    Set RS121 = Nothing
    DB120.Close
    Set DB120 = Nothing
    'List1.Enabled = False
    'Combo1.Enabled = False
    'txtsearch.Locked = True
End If
End Sub

Private Sub Timer1_Timer()
If (Label4.Left + Label4.Width) <= 0 Then
     Label4.Left = Me.Width
End If
Label4.Left = Label4.Left - 100
End Sub

Private Sub txtcustname_Change()
txtcustname.BackColor = vbWhite
End Sub

Private Sub txtdiscount_Change()
txtamount = (Val(txtprice) * Val(txtquantity)) - (Val(txtdiscount))
End Sub

Private Sub txtdiscount_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
    Exit Sub
Else
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub txtquantity_Change()
If Val(txtprice.Text) = 0 Then
    txtquantity.Text = "0"
    MsgBox "Please update the price of Goods ..price must not be zero..please verify", vbInformation, "WRONG PRICE OF GOODS"
    MsgBox "Please use THE UPDATE RECORD FORM to correct the price of goods", , "VERIFY"
    txtprice.SetFocus
    txtprice.SelStart = 0
    txtprice.SelLength = Len(txtquantity.Text)
    Exit Sub
Else
    txtquantity.BackColor = vbWhite
    txtamount = (Val(txtprice) * Val(txtquantity)) - (Val(txtdiscount))
End If
End Sub

Private Sub txtquantity_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
    Exit Sub
Else
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub txtsearch_Change()
List1.clear
Dim DB11 As New ADODB.CONNECTION
Dim RS119 As New ADODB.Recordset
If DB11.State = adStateOpen Then DB11.Close
    DB11.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
    'RS11.Open "SELECT * FROM master where Description LIKE '%" & Trim(txtsearch.Text) & "%'", DB11, 3, 3
    
    RS119.Open "SELECT * FROM master where Description LIKE '%" & Trim(txtsearch.Text) & "%' AND unit = '" & Combo1.Text & "'", DB11, 3, 3
Do Until RS119.EOF
    List1.AddItem RS119!Description
    RS119.MoveNext
Loop
RS119.Close
Set RS119 = Nothing
DB11.Close
Set DB11 = Nothing
End Sub

Private Sub txtsearch2_Change()

''''CLEAR THE LIST '''''''
List2.clear
List3.clear
List4.clear
Dim DB11 As New ADODB.CONNECTION
Dim RS11 As New ADODB.Recordset
DB11.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS11.Open "SELECT * FROM TEMPSOLD where Description LIKE '%" & Trim(txtsearch2.Text) & "%'", DB11, 3, 3

'RS11.Open "SELECT * FROM master where Description LIKE '%" & Trim(txtsearch.Text) & "%' AND unit = '" & Combo1.Text & "'", DB11, 3, 3

Do Until RS11.EOF
    List2.AddItem RS11!Description
    List3.AddItem RS11!Quantity
    List4.AddItem RS11!amount
    RS11.MoveNext
Loop
RS11.Close
Set RS11 = Nothing
DB11.Close
Set DB11 = Nothing
End Sub
