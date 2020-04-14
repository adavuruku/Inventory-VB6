VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "PROGRAM LISTING"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   15855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C000&
      Caption         =   "PROGRAM_SEVENTEEN CONSTANTS"
      ForeColor       =   &H00400000&
      Height          =   2295
      Left            =   11160
      TabIndex        =   37
      Top             =   5280
      Width           =   4575
      Begin VB.TextBox txtQ7 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2760
         TabIndex        =   42
         Text            =   "0.5"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtoilfriction 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2760
         TabIndex        =   40
         Text            =   "0.03"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtdrivetrainefficiency 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2760
         TabIndex        =   38
         Text            =   "0.93"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Rolling Arms Length"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Oil Friction"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1320
         TabIndex        =   41
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Drive Train Efficiency"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C000&
      Caption         =   "PROGRAM_SIXTEEN CONSTANTS"
      ForeColor       =   &H00400000&
      Height          =   975
      Left            =   11160
      TabIndex        =   34
      Top             =   4080
      Width           =   4575
      Begin VB.TextBox txtq16 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2760
         TabIndex        =   35
         Text            =   "0.5"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Rolling Arms Length"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      Caption         =   "PROGRAM_FIFTEEN CONSTANTS"
      ForeColor       =   &H00400000&
      Height          =   1095
      Left            =   11160
      TabIndex        =   29
      Top             =   2760
      Width           =   4575
      Begin VB.TextBox txtmn 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   840
         TabIndex        =   31
         Text            =   "1.2"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtc 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2880
         TabIndex        =   30
         Text            =   "0.37"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "%Mn"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "%C"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2280
         TabIndex        =   32
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "PROGRAM_FOURTEEN CONSTANTS"
      ForeColor       =   &H00400000&
      Height          =   1095
      Left            =   11160
      TabIndex        =   26
      Top             =   1440
      Width           =   4575
      Begin VB.TextBox txttemperature 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2640
         TabIndex        =   27
         Text            =   "1250"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TEMPERATURE"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   600
         TabIndex        =   28
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "PROGRAM_NINE CONSTANTS"
      ForeColor       =   &H00400000&
      Height          =   2655
      Left            =   3120
      TabIndex        =   19
      Top             =   5160
      Width           =   3975
      Begin VB.TextBox txtfrictionalcontactarea 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Text            =   "0.975"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox txtcontactfriction 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Text            =   "0.8"
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Frictional Contact Area"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Friction"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "PROGRAM_THIRTEEN CONSTANTS"
      ForeColor       =   &H00400000&
      Height          =   1095
      Left            =   11160
      TabIndex        =   18
      Top             =   240
      Width           =   4575
      Begin VB.TextBox txtspeed 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2640
         TabIndex        =   25
         Text            =   "15"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "SPEED M/S"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1200
         TabIndex        =   24
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command22 
      Caption         =   "PROGRAM_SEVENTEEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   14
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton Command21 
      Caption         =   "PROGRAM_SIXTEEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   13
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton Command20 
      Caption         =   "PROGRAM_FIFTEEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   12
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton Command19 
      Caption         =   "PROGRAM_FOURTEEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   11
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton Command18 
      Caption         =   "PROGRAM_THIRTEEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   10
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton Command17 
      Caption         =   "PROGRAM_TWELVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   9
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton Command16 
      Caption         =   "PROGRAM_ELEVEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   8
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command15 
      Caption         =   "PROGRAM_TEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command11 
      Caption         =   "PROGRAM_SIX _TO_NINE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   6
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton Command10 
      Caption         =   "PROGRAM_FIVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Caption         =   "PROGRAM_FOUR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      Caption         =   "PROGRAM_THREE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PROGRAM_TWO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PROGRAM_ONE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   13440
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   11040
      TabIndex        =   17
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   7560
      TabIndex        =   16
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   2880
      TabIndex        =   15
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command10_Click()
Dim DB2 As New ADODB.Connection
Dim RS3 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
p = 12
t = 1
If DB2.State = adStateOpen Then DB2.Close
DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
If RS3.State = adStateOpen Then RS3.Close
RS3.Open "SELECT * FROM Maths2", DB2, 3, 3
RS3.MoveLast
Do Until RS3.BOF
    
    If t <> 1 Then
       ' MsgBox RS3!Elongation
        'update 2
        If RS2.State = adStateOpen Then RS2.Close
        RS2.Open "SELECT * FROM maths2 where stand = '" & t & "'", DB2, 3, 3
        
        felong = RS3!Elongation
        j = Val(j) * Val(felong)
        
        RS2!lenght = j
        
        RS2.Update
    Else
        j = 12

        j = Val(j)

     ' If RS3.State = adStateOpen Then RS3.Close
    '  RS3.Open "SELECT * FROM Maths where Stand = '" & t & "'", DB2, 3, 3
      
      If RS2.State = adStateOpen Then RS2.Close
      RS2.Open "SELECT * FROM Maths2 where Stand = '" & t & "'", DB2, 3, 3
      RS2!lenght = j * Val(RS3!Elongation)
      j = j * Val(RS3!Elongation)
      RS2.Update
      
      farea = j
      felong = RS3!Elongation

    End If
t = Val(t) + 1
RS3.MovePrevious
Loop
message ("PROGRAM FIVE")
End Sub

Private Sub Command11_Click()
If ((txtcontactfriction.Text = "") Or (txtfrictionalcontactarea.Text = "")) Then
    MsgBox "The constant fields are empty..they require a valid number or zero"
    txtcontactfriction.SetFocus
    Exit Sub
Else

Dim DB2 As New ADODB.Connection
Dim RS3 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim RS5 As New ADODB.Recordset

Dim d19 As Double

If DB2.State = adStateOpen Then DB2.Close
DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
If RS3.State = adStateOpen Then RS3.Close
RS3.Open "SELECT * FROM Maths2", DB2, 3, 3
H0 = 120
p = 19

'ENTRANCE OF 19 CONSTANT WILL B CHANGING
Entrance_b = 12.01
RS3.MoveFirst
Do Until RS3.EOF
num1 = p / 2 'used to get the even and odd
If p = 1 Then
    'is a constant
    Entrance_h = 120
    Entrance_b = 120
    
    If RS2.State = adStateOpen Then RS2.Close
    RS2.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB2, 3, 3
   
    pico = Val(txtcontactfriction.Text) * (1.05 - (0.0005 * (1250)))
    pico = pico * pico
    deltah = Val(txtcontactfriction.Text) * pico * 250
    
    Exit_h = Entrance_h - deltah
       
    ' calculating the absolute spread
    
    'change in small b and big B
    Chnage_b = 0.4 * (deltah / 120) * ((250 * deltah) ^ 0.5)
    actual_spread = 120 + Chnage_b
    
    Exit_b = actual_spread
    
    'big B
    BValue = actual_spread + 12
    
    'calculating Area
    'A1=b1 x h1 x f
    f = Val(txtfrictionalcontactarea.Text)
    Area = actual_spread * Exit_h * f
    
    'open database for saving the values two recordset rs2 is at 1 and rs4 is at 2
    h = p + 1
    If RS4.State = adStateOpen Then RS4.Close
    RS4.Open "SELECT * FROM maths2 where stand = '" & h & "'", DB2, 3, 3
    
    'record as at stand one
    RS2!Entrance_h = Val(Entrance_h)
    RS2!Entrance_b = Val(Entrance_b)
    RS2!Exit_h = Val(Exit_h)
    RS2!Exit_b = Val(Exit_b)
    RS2!b_Value = Val(actual_spread)
    RS2!h_Value = Val(Exit_h)
    RS2!BValue = Val(BValue)
    RS2!Delta_H = Val(deltah)
    RS2!Area = Val(Area)
    
    'record as at stand 2
    RS4!Entrance_h = Val(Exit_h)
    RS4!Entrance_b = Val(Exit_b)

        RS2.Update
        RS2.Close
        Set RS2 = Nothing
    RS4.Update
    RS4.Close
    Set RS4 = Nothing
   ' MsgBox p & Space(4) & Exit_h & Space(4) & deltah & Space(4) & BValue & Space(4) & b_Value
    
ElseIf p = 2 Then
    tt = p - 1
    If RS5.State = adStateOpen Then RS5.Close
    RS5.Open "SELECT * FROM maths2 where stand = '" & tt & "'", DB2, 3, 3
    
    'retrieve exit of previous(1) for next as entrance
    Entrance_h = RS5!Exit_h
    Entrance_b = RS5!Exit_b
    
    'calculating change in small b and big B
    pico = Val(txtcontactfriction.Text) * (1.05 - (0.0005 * (1225)))
    pico = pico * pico
    deltah = Val(txtcontactfriction.Text) * pico * 250
    Exit_h = Entrance_h - deltah
    
    Change_b = 0.4 * (deltah / Val(Entrance_h) * ((250 * deltah) ^ 0.5))
    
    'calculate the BValue and small b_Value
    actual_spread = Entrance_b + Change_b
    Exit_b = Val(actual_spread)
    BigB = actual_spread + 12
    
    'calculate the Area
    f = Val(txtfrictionalcontactarea.Text)
    Area = actual_spread * Exit_h * f
    
    'save all the record rs2 is for spread as at (2) rs5 is for spread as at (1) rs4 is for spread as at (3)
    If RS2.State = adStateOpen Then RS2.Close
    RS2.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB2, 3, 3
    
    h = p + 1
    If RS4.State = adStateOpen Then RS4.Close
    RS4.Open "SELECT * FROM maths2 where stand = '" & h & "'", DB2, 3, 3
    
    'record at stand 2
    RS2!b_Value = Val(actual_spread)
    RS2!h_Value = Val(Exit_h)
    RS2!BValue = BigB
    RS2!Area = Area
    RS2!Exit_h = Val(Exit_h)
    RS2!Exit_b = Val(Exit_b)
    RS2!Delta_H = Val(deltah)
    RS2.Update
    RS2.Close
    Set RS2 = Nothing
    
    'record at stand 3
    RS4!Entrance_h = Val(Exit_h)
    RS4!Entrance_b = Val(Exit_b)
    RS4.Update
    RS4.Close
    Set RS4 = Nothing
ElseIf p = 19 Then
        'there entrance and exit h
      ' tt = p - 1
      ' If RS4.State = adStateOpen Then RS4.Close
        'RS4.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB2, 3, 3
        
        si = Sin(45)
        co = Cos(45)
        b = 12.1
        b2 = b * b
        S = 1.6
        s2 = S * S
       ' d19 = 12.01
        d19 = 12.02
        d2 = d * d
        r = 6.01
        r2 = r * r

        'up value
        fu = (b2 + s2 + (4 * r2)) - (4 * r * ((S * si) + (b * co)))
        'downvalue
        fd3 = 8 * r - (4 * (S * si + b * co))

        'final resulr
        final = fu / fd3
        'MsgBox final
        If RS2.State = adStateOpen Then RS2.Close
        RS2.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB2, 3, 3
        
        RS2!Radius = final
        RS2!BValue = b
    
        RS2!b_Value = d19
        
        'THIS IS FOR H
        RS2!Exit_h = d19
      ' RS2!Exit_h = 12.02
        RS2!Entrance_h = Entrance_h
        Entrance_h = Entrance_h
        
        'THIS IS FOR B
        RS2!Exit_b = Val(d19)
        RS2!Entrance_b = Val(d19)
        
        RS2!h_Value = d19
        RS2.Update
        RS2.Close
        Set RS2 = Nothing
        'STOP HERE
 'second opt (even)
ElseIf (num1 = Int(num1) And (p <> 1 And p <> 2 And p <> 19)) Then 'even numbers
     '  k2 = 0.95 'constant
        
        'k2 = 1.2 'constant
       ' h18 = k2 * d19
    
    'search for area and constants (both two constatnts) using the value
        If RS2.State = adStateOpen Then RS2.Close
        RS2.Open "SELECT * FROM Maths2 where Stand = '" & p & "'", DB2, 3, 3
         j = Val(RS2!Area)
         S18 = Val(RS2!SValue)
         
         k2 = Val(RS2!constant) 'constant k2
         
         h18 = k2 * d19 'note d19 rep  bn
            
            'MsgBox "SValue = " & S18
            'k18 = 0.45 'constant
            
            k18 = Val(RS2!constanteven) 'constant Kn
            m18 = k18 * h18
             
            'THE SPREAD AT stand Kn(even)
            b18 = (3 * j) / (h18 * (2 + k18))
            
            'm18 = k18 * h18
            
            hhh = b18 * b18
            r100 = (h18 - m18) ^ 2
            r1001 = hhh + r100
            r100d = 4 * (h18 - m18)
            ans = r1001 / r100d
            
            'calculate the big B Value
            xa = (h18 - S18) * ans
            xa2 = (h18 - S18) / 2
            xa3 = xa2 ^ 2
            xa4 = xa - xa3
            finalB = 2 * ((xa4) ^ 0.5)
        RS2.Close
        Set RS2 = Nothing
            
            'update the record to db
           
            If RS2.State = adStateOpen Then RS2.Close
            RS2.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB2, 3, 3
             
             h = p + 1
            If RS4.State = adStateOpen Then RS4.Close
            RS4.Open "SELECT * FROM maths2 where stand = '" & h & "'", DB2, 3, 3
            
            RS2!Radius = ans
            RS2!BValue = finalB
           ' RS2!b_Value = d19
            RS2!b_Value = b18
            RS2!h_Value = h18
            
            'THIS IS FOR H
            'RS2!Exit_h = Entrance_h
            RS2!Exit_h = h18
            RS4!Entrance_h = h18
            Entrance_h = h18
            
            'THIS IS FOR B
            RS2!Exit_b = Val(b18)
            RS4!Entrance_b = Val(b18)
            
            'RS2!Exit_b = Entrance_b
            'RS2!Entrance_b = d19
           ' Entrance_b = d19
            RS2.Update
            RS2.Close
            Set RS2 = Nothing
                 RS4.Update
                RS4.Close
                Set RS4 = Nothing
            d19 = 0 'u will put new value here from ODD         'ending even
            
            'second opt (odd)
ElseIf (num1 <> Int(num1) And (p <> 1 And p <> 2 And p <> 19)) Then 'odd numbers
            'k1 = 1.27 'constant
            'search for area using the value
            If RS2.State = adStateOpen Then RS2.Close
            RS2.Open "SELECT * FROM Maths2 where Stand = '" & p & "'", DB2, 3, 3
            j = Val(RS2!Area)
            
            K1 = Val(RS2!constant) 'constant k1
            
            bg1 = 24 * K1 * j
            bg2 = 3 + (16 * (K1 * K1))
            bg3 = bg1 / bg2
            d19 = bg3 ^ 0.5 'note d19 rep  bn
            
            'height of the stock
            h17 = K1 * d19
            
            'radius of the groove
            r17a = (h17 * h17) + (d19 * d19)
            r17b = 4 * d19
            finalr17 = r17a / r17b
            
            'capital B
            finalB = d19 + 1
            
        RS2.Close
        Set RS2 = Nothing

        'update the record to db
        If RS2.State = adStateOpen Then RS2.Close
            RS2.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB2, 3, 3
            
             h = p + 1
            If RS4.State = adStateOpen Then RS4.Close
            RS4.Open "SELECT * FROM maths2 where stand = '" & h & "'", DB2, 3, 3
            
            RS2!Radius = finalr17
            RS2!BValue = finalB
            RS2!b_Value = d19
            RS2!h_Value = h17
            
            'THIS IS FOR H
            'rs2 is for the current value
            RS2!Exit_h = h17
            'rs4 points to the value 1 ahead
            RS4!Entrance_h = h17
            Entrance_h = h17
            
            'THIS IS FOR B
            RS2!Exit_b = Val(d19)
            RS4!Entrance_b = Val(d19)
            Entrance_b = Val(d19)
            
            'RS2!Exit_b = Entrance_b
            'RS2!Entrance_b = d19
            
            'If p = 3 Then
            'MsgBox "is 3 " & Entrance_b
            'End If
            
            RS2.Update
            RS2.Close
            Set RS2 = Nothing
            RS4.Update
            RS4.Close
            Set RS4 = Nothing
Else
End If
p = p - 1

RS3.MoveNext
Loop
End If
message ("PROGRAM SIX - NINE")
End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command13_Click()

End Sub

Private Sub Command14_Click()

End Sub

Private Sub Command15_Click()
Dim DB3 As New ADODB.Connection
Dim RS11 As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim RS14 As New ADODB.Recordset

If DB3.State = adStateOpen Then DB3.Close
DB3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
If RS11.State = adStateOpen Then RS11.Close
RS11.Open "SELECT Entrance_h, Exit_h FROM Maths2", DB3, 3, 3
p = 19
RS11.MoveFirst
Do Until RS11.EOF

    If (Val(RS11!Exit_h) > Val(RS11!Entrance_h)) Then
        q = p - 1
        'retrieve b_Value as at 18
        If RS12.State = adStateOpen Then RS12.Close
        RS12.Open "SELECT stand, b_Value FROM Maths2 where stand='" & q & "'", DB3, 3, 3
            deduct_val = 0
            deduct_val = Val(RS12!b_Value) - Val(RS11!Exit_h)
        RS12.Close
        Set RS12 = Nothing
    Else
        deduct_val = 0
        deduct_val = Val(RS11!Entrance_h) - Val(RS11!Exit_h)
    End If
'update here
If RS14.State = adStateOpen Then RS14.Close
    RS14.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB3, 3, 3
        RS14!Delta_H = Val(deduct_val)
    RS14.Update
    RS14.Close
    Set RS14 = Nothing
'decrease p for next loop
p = p - 1
If p = 2 Then
Exit Do
End If
RS11.MoveNext
Loop
'close initial database
RS11.Close
Set RS11 = Nothing
message ("PROGRAM TEN")
End Sub

Private Sub Command16_Click()
Dim DB3 As New ADODB.Connection
Dim RS11 As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim RS14 As New ADODB.Recordset

If DB3.State = adStateOpen Then DB3.Close
DB3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
If RS11.State = adStateOpen Then RS11.Close
RS11.Open "SELECT * FROM Maths2", DB3, 3, 3
p = 19
RS11.MoveFirst
Do Until RS11.EOF

Dw = (RS11!Diameter_roll - (RS11!Area / RS11!b_Value)) + RS11!SValue

Rw = Val(Val(Dw) / 2)

angle_of_bite = (Val(RS11!Delta_H) / Val(Rw)) ^ 0.5

total_angle_of_bite_degree = 57.5 * Val(angle_of_bite)

'update here
If RS14.State = adStateOpen Then RS14.Close
    RS14.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB3, 3, 3
        RS14!Effective_Diameter = Val(Dw)
        RS14!Effective_Radius = Val(Rw)
        RS14!angle_of_bite = Val(angle_of_bite)
        RS14!angle_of_bite_deg = Val(total_angle_of_bite_degree)
    RS14.Update
    RS14.Close
    Set RS14 = Nothing
'decrease p for next loop

p = p - 1

'use this to sum up all the angle of bite in other to use the value for the degree
'total_angle_of_bite = Val(total_angle_of_bite) + Val(angle_of_bite)

RS11.MoveNext
Loop
RS11.Close
Set RS11 = Nothing
message ("PROGRAM ELLEVEN")
End Sub

Private Sub Command17_Click()
Dim DB3 As New ADODB.Connection
Dim RS11 As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim RS14 As New ADODB.Recordset

If DB3.State = adStateOpen Then DB3.Close
DB3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
If RS11.State = adStateOpen Then RS11.Close
RS11.Open "SELECT * FROM Maths2", DB3, 3, 3
p = 19
RS11.MoveFirst
Do Until RS11.EOF

Dw = (RS11!Diameter_roll - (RS11!Area / RS11!b_Value)) + RS11!SValue

Rw = Val(Val(Dw) / 2)

angle_of_bite = (Val(RS11!Delta_H) / Val(Rw)) ^ 0.5

total_angle_of_bite_degree = 57.5 * Val(angle_of_bite)

'update here
If RS14.State = adStateOpen Then RS14.Close
    RS14.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB3, 3, 3
        RS14!Effective_Diameter = Val(Dw)
        RS14!Effective_Radius = Val(Rw)
        RS14!angle_of_bite = Val(angle_of_bite)
        RS14!angle_of_bite_deg = Val(total_angle_of_bite_degree)
    RS14.Update
    RS14.Close
    Set RS14 = Nothing
'decrease p for next loop

p = p - 1

'use this to sum up all the angle of bite in other to use the value for the degree
'total_angle_of_bite = Val(total_angle_of_bite) + Val(angle_of_bite)

RS11.MoveNext
Loop
RS11.Close
Set RS11 = Nothing
message ("PROGRAM TWELVE")
End Sub

Private Sub Command18_Click()

Dim DB3 As New ADODB.Connection
Dim RS11 As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim RS14 As New ADODB.Recordset

If (txtspeed.Text = "") Then
    MsgBox "The constant fields are empty..they require a valid number or zero"
    txtspeed.SetFocus
    Exit Sub
Else

    If DB3.State = adStateOpen Then DB3.Close
    DB3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
        j = 19
        If RS14.State = adStateOpen Then RS14.Close
        RS14.Open "SELECT * FROM maths2 where stand = '" & j & "'", DB3, 3, 3
        
            pile = 22 / 7
            Vr = Val(txtspeed.Text)
            Nr = ((Vr * 60 * 1000) / (pile * Val(RS14!Effective_Diameter)))
            
            Nm = Val(Nr) * Val(RS14!Transmission_Ratio)
            
            Q_K = (Val(RS14!Effective_Diameter) * Val(Nr) * Val(RS14!Area))
        
            'save values for stand 19
            RS14!moto_revolution_Nm = Val(Nm)
            RS14!revolution_of_roll_Nm = Val(Nr)
            RS14!rolling_constant = Val(Q_K)
            RS14!speed = Val(Vr)
        RS14.Update
        RS14.Close
        Set RS14 = Nothing
    If RS11.State = adStateOpen Then RS11.Close
    RS11.Open "SELECT * FROM Maths2", DB3, 3, 3
    p = 19
    RS11.MoveFirst
    Do Until RS11.EOF
        If (RS11!Stand <> 19) Then
                pile = 22 / 7
                
                Nr = (Val(Q_K) / (Val(RS11!Effective_Diameter) * Val(RS11!Area)))
            
                Nm = Val(Nr) * Val(RS11!Transmission_Ratio)
                
                Vr = ((Val(Nr) * pile * Val(RS11!Effective_Diameter)) / (60000))
        
                'update here
                If RS14.State = adStateOpen Then RS14.Close
                RS14.Open "SELECT * FROM maths2 where stand = '" & p & "'", DB3, 3, 3
                    RS14!moto_revolution_Nm = Val(Nm)
                    RS14!revolution_of_roll_Nm = Val(Nr)
                    RS14!rolling_constant = Val(Q_K)
                    RS14!speed = Val(Vr)
                RS14.Update
                RS14.Close
                Set RS14 = Nothing
        
        
        End If
    
    'decrease p for next loop
    p = p - 1
    
    RS11.MoveNext
    Loop
    RS11.Close
    Set RS11 = Nothing
End If
message ("PROGRAM THIRTEEN")
End Sub

Private Sub Command19_Click()
Dim DB3 As New ADODB.Connection
Dim RS11 As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim RS14 As New ADODB.Recordset

If (txttemperature.Text = "") Then
    MsgBox "The constant fields are empty..they require a valid number or zero"
    txttemperature.SetFocus
    Exit Sub
Else

    If DB3.State = adStateOpen Then DB3.Close
    DB3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
    
    'retrieve height at 1 and 19
    If RS11.State = adStateOpen Then RS11.Close
    RS11.Open "SELECT Stand, h_Value FROM Maths2", DB3, 3, 3
    p = 19
    RS11.MoveFirst
    Do Until RS11.EOF
        If (RS11!Stand = 19) Then
            Hf = RS11!h_Value
        End If
        
        If (RS11!Stand = 1) Then
            Ho = RS11!h_Value
        End If
    RS11.MoveNext
    Loop
    RS11.Close
    Set RS11 = Nothing
    'retrieve height at 1 and 19 ends here
    
    'calculate all the value ..default and update at 19
        T1 = Val(txttemperature.Text)
        Tf = 800
        K1 = ((T1 - Tf) / (Ho - Hf)) * Ho * Hf
        Too = (T1 + (Val(K1) / Val(Ho)))
        Tn = (Too - (Val(K1) / Val(Hf)))
        
        'Save all at 19
        j = 19
       ' MsgBox j & " " & Too & " " & K1
        If RS14.State = adStateOpen Then RS14.Close
        RS14.Open "SELECT * FROM maths2 where stand = '" & j & "'", DB3, 3, 3
            RS14!Temperature = Tn
        RS14.Update
        RS14.Close
        Set RS14 = Nothing
    'all default ND Sve at 19 ends here
    
    'save for stand 18 to stand 1
    If RS11.State = adStateOpen Then RS11.Close
    RS11.Open "SELECT * FROM Maths2", DB3, 3, 3
    q = 19
    RS11.MoveFirst
    Do Until RS11.EOF
        If (RS11!Stand <> 19) Then
                 Tn = (Too - (Val(K1) / Val(RS11!h_Value)))
                ' MsgBox q & " " & Too & " " & K1
                'update here
                If RS14.State = adStateOpen Then RS14.Close
                RS14.Open "SELECT * FROM maths2 where stand = '" & q & "'", DB3, 3, 3
                     RS14!Temperature = Tn
                RS14.Update
                RS14.Close
                Set RS14 = Nothing
        End If
    
    'decrease q for next loop
    q = q - 1
    
    RS11.MoveNext
    Loop
    RS11.Close
    Set RS11 = Nothing
End If
message ("PROGRAM FOURTEEN")
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command20_Click()
Dim DB3 As New ADODB.Connection
Dim RS11 As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim RS14 As New ADODB.Recordset
'P = Pm * Ac
'constant for all stand
If ((txtmn.Text = "") Or (txtc.Text = "")) Then
    MsgBox "The constant fields are empty..they require a valid number or zero"
    txtmn.SetFocus
    Exit Sub
Else

    C = Val(txtc.Text)
    Mn = Val(txtmn.Text)
    'calculate Pm
    If DB3.State = adStateOpen Then DB3.Close
    DB3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
    
    'retrieve height at 1 and 19
    If RS11.State = adStateOpen Then RS11.Close
    RS11.Open "SELECT * FROM Maths2", DB3, 3, 3
    p = 19
    RS11.MoveFirst
    Do Until RS11.EOF
        
        'determine even and odd from the loop
        geteven_od = Val(RS11!Stand) / 2
        
        'do oddd
        If (Val(geteven_od) <> Int(Val(geteven_od)) And (Val(RS11!Stand) <> 1)) Then
                
                'pick backward (previous)  values
                vallprev = Val(RS11!Stand) - 1
                vallnext = Val(RS11!Stand) + 1
                
                'pick value previous
                If RS14.State = adStateOpen Then RS14.Close
                RS14.Open "SELECT * FROM maths2 where stand = '" & vallprev & "'", DB3, 3, 3
                    Bprev = RS14!b_Value
                    Hprev = RS14!h_Value
                RS14.Close
                Set RS14 = Nothing
            
            'do all the calculation for even
            Zigma = (14 - (0.01 * RS11!Temperature)) * (1.4 + C + Mn)
            elongation_n = 0.01 * (14 - (0.01 * RS11!Temperature))
            coefficient_friction = (1.05 - (0.0005 * RS11!Temperature))
            
            'convert values to meter before using it for slant v
            delta_hconvert = Val(RS11!Delta_H) / 1000
            Bprev_convert = Val(Bprev) / 1000
            h_standconvert = Val(RS11!h_Value) / 1000
            Rw_convert = Val(RS11!Effective_Radius) / 1000
            
            'calculation of slanting v begin
            sqroot = (delta_hconvert / Rw_convert) ^ 0.5
            upvalue = 2 * Val(RS11!speed) * sqroot
            downvalue = Bprev_convert + h_standconvert
            'so therefore
            v = Val(upvalue / downvalue)
            
            'calculation of Kf begin
            'step1 = (1.6 * coefficient_friction * (Val(RS11!Delta_H) / Val(RS11!Effective_Radius)) ^ 0.5) - (1.2 * Val(RS11!Delta_H))
            step1 = (1.6 * coefficient_friction * (Val(RS11!Delta_H) * Val(RS11!Effective_Radius)) ^ 0.5) - (1.2 * Val(RS11!Delta_H))
            step2 = step1 / (Val(RS11!h_Value) + Val(Bprev))
            Kf = 1 + step2
            
            'calculation of Pm
            Pm = Kf * (Zigma + (v * elongation_n))
            
            'calcilating of Ac
            Acstep1 = (Val(RS11!Delta_H) * Val(RS11!Effective_Radius)) ^ 0.5
            Acstep2 = (Bprev + Val(RS11!BValue)) / 2
            Ac = Acstep1 * Acstep2
            
            'calculation of P
            Preal = Val(Pm) * Val(Ac)
            'p = (Preal / 10000) / 9.81
            p = (Preal * 9.81) / 1000000
            'save all the calculated value
            If RS14.State = adStateOpen Then RS14.Close
            RS14.Open "SELECT * FROM maths2 where stand = '" & RS11!Stand & "'", DB3, 3, 3
                RS14!Rolling_Pressure = Val(Pm)
                RS14!Rolling_Load_KgF = Val(Preal)
                RS14!Rolling_Load_MN = Val(p)
                RS14!Contact_Area = Val(Ac)
            RS14.Update
            RS14.Close
            Set RS14 = Nothing
             'MsgBox "Am  odd at " & RS11!Stand
        End If
        
        'do even and Stand One since they are thesame
        If (Val(geteven_od) = Int(Val(geteven_od)) Or (Val(RS11!Stand) = 1)) Then
            
            'pick backward (previous)  values
            vallprev = Val(RS11!Stand) - 1
            vallnext = Val(RS11!Stand) + 1
            If (Val(RS11!Stand) = 1) Then
                Bprev = 120
                Hprev = 120
                Bnext = 120
                Hnext = 120
            Else
                'pick value previous
                If RS14.State = adStateOpen Then RS14.Close
                RS14.Open "SELECT * FROM maths2 where stand = '" & vallprev & "'", DB3, 3, 3
                    'Bprev = RS14!b_Value
                    Bprev = RS14!BValue
                    Hprev = RS14!h_Value
                RS14.Close
                Set RS14 = Nothing
                
                'pick value next
                If RS14.State = adStateOpen Then RS14.Close
                RS14.Open "SELECT * FROM maths2 where stand = '" & vallnext & "'", DB3, 3, 3
                    Bnext = RS14!BValue
                    Hnext = RS14!h_Value
                RS14.Close
                Set RS14 = Nothing
            End If
            'picking ends here
            
            'do all the calculation for even
            Zigma = (14 - (0.01 * RS11!Temperature)) * (1.4 + C + Mn)
            elongation_n = 0.01 * (14 - (0.01 * RS11!Temperature))
            coefficient_friction = (1.05 - (0.0005 * RS11!Temperature))
            
            'convert values to meter before using it for slant v
            delta_hconvert = Val(RS11!Delta_H) / 1000
            Hprev_convert = Val(Hprev) / 1000
            h_standconvert = Val(RS11!h_Value) / 1000
            Rw_convert = Val(RS11!Effective_Radius) / 1000
            
            'calculation of slanting v begin
            sqroot = (delta_hconvert / Rw_convert) ^ 0.5
            upvalue = 2 * Val(RS11!speed) * sqroot
            downvalue = Hprev_convert + h_standconvert
            'so therefore
            v = Val(upvalue / downvalue)
            
            'calculation of Kf begin
            'step1 = (1.6 * coefficient_friction * (Val(RS11!Delta_H) / Val(RS11!Effective_Radius)) ^ 0.5) - (1.2 * Val(RS11!Delta_H))
            step1 = (1.6 * coefficient_friction * (Val(RS11!Delta_H) * Val(RS11!Effective_Radius)) ^ 0.5) - (1.2 * Val(RS11!Delta_H))
            step2 = step1 / (Val(RS11!h_Value) + Val(Hnext))
            Kf = 1 + step2
            
            'calculation of Pm
            Pm = Kf * (Zigma + (v * elongation_n))
            
            'calcilating of Ac
            Acstep1 = (Val(RS11!Delta_H) * Val(RS11!Effective_Radius)) ^ 0.5
            Acstep2 = (Bprev + Val(RS11!BValue)) / 2
            Ac = Acstep1 * Acstep2
            
            'calculation of P
            Preal = Val(Pm) * Val(Ac)
             'p = (Preal / 10000) / 9.81
             p = (Preal * 9.81) / 1000000
            'save all the calculated value
            If RS14.State = adStateOpen Then RS14.Close
            RS14.Open "SELECT * FROM maths2 where stand = '" & RS11!Stand & "'", DB3, 3, 3
                RS14!Rolling_Pressure = Val(Pm)
                RS14!Rolling_Load_KgF = Val(Preal)
                RS14!Rolling_Load_MN = Val(p)
                RS14!Contact_Area = Val(Ac)
            RS14.Update
            RS14.Close
            Set RS14 = Nothing
        End If
    RS11.MoveNext
    Loop
    RS11.Close
    Set RS11 = Nothing
End If
message ("PROGRAM FIFTEEN")
End Sub

Private Sub Command21_Click()
Dim DB3 As New ADODB.Connection
Dim RS11 As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim RS14 As New ADODB.Recordset

If (txtq16.Text = "") Then
    MsgBox "The constant fields are empty..they require a valid number or zero"
    txtq16.SetFocus
    Exit Sub
Else
    If DB3.State = adStateOpen Then DB3.Close
    DB3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
    
    If RS11.State = adStateOpen Then RS11.Close
    RS11.Open "SELECT * FROM Maths2", DB3, 3, 3
    'p = 19
    RS11.MoveFirst
    Do Until RS11.EOF
        constant = Val(txtq16.Text)
        pile = 22 / 7
        
        'calculate
        Work_Done = 4 * Val(RS11!Rolling_Load_MN) * pile * constant * Val((Val(RS11!Effective_Radius) / 1000)) * Val(RS11!angle_of_bite)
        
        'save all the calculated value
            If RS14.State = adStateOpen Then RS14.Close
            RS14.Open "SELECT * FROM maths2 where stand = '" & RS11!Stand & "'", DB3, 3, 3
                RS14!Work_Done = Val(Work_Done)
            RS14.Update
            RS14.Close
            Set RS14 = Nothing
    
    RS11.MoveNext
    Loop
    RS11.Close
    Set RS11 = Nothing
End If
message ("PROGRAM SIXTEEN")
End Sub

Private Sub Command22_Click()
Dim DB3 As New ADODB.Connection
Dim RS11 As New ADODB.Recordset
Dim RS12 As New ADODB.Recordset
Dim RS14 As New ADODB.Recordset

If ((txtdrivetrainefficiency.Text = "") Or (txtoilfriction.Text = "") Or (txtQ7.Text = "")) Then
    MsgBox "The constant fields are empty..they require a valid number or zero"
    txtdrivetrainefficiency.SetFocus
    Exit Sub
Else
    If DB3.State = adStateOpen Then DB3.Close
    DB3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
    
    If RS11.State = adStateOpen Then RS11.Close
    RS11.Open "SELECT * FROM Maths2", DB3, 3, 3
    'p = 19
    RS11.MoveFirst
    Do Until RS11.EOF
        constant = Val(txtQ7.Text)
        pile = 22 / 7
        
        'calculate
          
        'Dynamic_Torgue calculation
        alpha = constant * Val((Val(RS11!Effective_Radius) / 1000)) * Val(RS11!angle_of_bite)
        Dynamic_Torgue = Val(RS11!Rolling_Load_KgF) * 2 * alpha * 0.001
        
        'Frictional_Torgue calculation
        constant_Q = Val(txtoilfriction.Text)
        Frictional_Torgue = Val(RS11!Rolling_Load_KgF) * constant_Q * Val(Val(RS11!D_Value) / 1000) * 0.001
        
        ' Rolling_Torgue calculation
        Rolling_Torgue = Val(Dynamic_Torgue) + Val(Frictional_Torgue)
        
        'Rolling_Power_hp calculation
        Rolling_Power_hp = ((Val(Rolling_Torgue) * Val(RS11!revolution_of_roll_Nm)) / 0.716)
        
        'Static_Power calculation
        Static_Power = ((Val(Rolling_Torgue) * Val(RS11!revolution_of_roll_Nm)) / Val(txtdrivetrainefficiency.Text)) * 1.025
        
        'save all the calculated value
            If RS14.State = adStateOpen Then RS14.Close
            RS14.Open "SELECT * FROM maths2 where stand = '" & RS11!Stand & "'", DB3, 3, 3
                
                RS14!Dynamic_Torgue = Val(Dynamic_Torgue)
                RS14!Frictional_Torgue = Val(Frictional_Torgue)
                RS14!Rolling_Torgue = Val(Rolling_Torgue)
                RS14!Rolling_Power_hp = Val(Rolling_Power_hp)
                RS14!Static_Power = Val(Static_Power)
              
            RS14.Update
            RS14.Close
            Set RS14 = Nothing
    
    RS11.MoveNext
    Loop
    RS11.Close
    Set RS11 = Nothing
    
message ("PROGRAM SEVENTEEN")
End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

'call the function(module) to perform the deduction
deduct
End Sub
Private Sub deduct()
effective_working_diameter
End Sub

Private Sub effective_working_diameter()
program_13
End Sub

Private Sub program_13()
program_14
End Sub
Private Sub program_14()
program_15
End Sub
Private Sub program_15()
program_16
End Sub
Private Sub program_16()
program_17

End Sub
Private Sub program_17()
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()
message ("PROGRAM ONE")
End Sub

Private Sub Command7_Click()
message ("PROGRAM TWO")
End Sub

Private Sub Command8_Click()
p = 113.47
t = 19

Dim DB2 As New ADODB.Connection
Dim RS3 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

If DB2.State = adStateOpen Then DB2.Close
DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
If RS3.State = adStateOpen Then RS3.Close
RS3.Open "SELECT * FROM Maths2", DB2, adOpenKeyset, adLockReadOnly
RS3.MoveFirst
Do Until RS3.EOF
If t = 19 Then
    j = 0
    'j = Val(RS3!Elongation) * Val(p)
    j = Val(p)
    'update 2
    If RS2.State = adStateOpen Then RS2.Close
    RS2.Open "SELECT * FROM Maths2 where Stand = '" & t & "'", DB2, 3, 3
    RS2!Area = j
    RS2.Update
    RS2.Close
    Set RS2 = Nothing
    farea = j
    felong = RS3!Elongation
ElseIf t = 1 Then
    If RS2.State = adStateOpen Then RS2.Close
    RS2.Open "SELECT * FROM Maths2 where Stand = '" & t & "'", DB2, 3, 3
    RS2!Area = 0
    RS2.Update
    RS2.Close
    Set RS2 = Nothing
ElseIf t = 2 Then
    If RS2.State = adStateOpen Then RS2.Close
    RS2.Open "SELECT * FROM Maths2 where Stand = '" & t & "'", DB2, 3, 3
    RS2!Area = 0
    RS2.Update
    RS2.Close
    Set RS2 = Nothing
Else
    j = 0
    j = Val(farea) * Val(felong)
    
    
    'update 2
    If RS2.State = adStateOpen Then RS2.Close
    RS2.Open "SELECT * FROM maths2 where stand = '" & t & "'", DB2, 3, 3
    RS2!Area = j
    'RS2!Area = 222
    RS2.Update
    RS2.Close
    Set RS2 = Nothing
    farea = j
    felong = RS3!Elongation
End If
t = Val(t) - 1
RS3.MoveNext
Loop
RS3.Close
Set RS3 = Nothing
message ("PROGRAM THREE")
End Sub

Private Sub Command9_Click()
    Dim DB2 As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim RS3 As New ADODB.Recordset
    
    If DB2.State = adStateOpen Then DB2.Close
    DB2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MATHS\maths.mdb;Persist Security Info=False"
    
     If RS3.State = adStateOpen Then rs.Close
        RS3.Open "SELECT * FROM Maths2", DB2, 3, 3
        
    ti = 19
    RS3.MoveFirst
    Do Until RS3.EOF
    
        tt = (ti) - 1
    
        If (tt = 0) Then
        'MsgBox "the real ti at ti = 1 tt = 0" & ti
            'it means is at stan one so use its default values
            A = 14400
            If rs.State = adStateOpen Then rs.Close
            rs.Open "SELECT * FROM Maths2 where Stand = '" & ti & "'", DB2, 3, 3
            rs!perc = Format(((Val(A) - Val(rs!Area)) / Val(rs!Area)) * 100, "###,####.00")
            
            'rs!perc = ((Val(A) - Val(rs!Area)) / Val(rs!Area)) * 100
            rs.Update
        Else
       ' MsgBox "the real ti" & ti
            If rs.State = adStateOpen Then rs.Close
            rs.Open "SELECT * FROM Maths2 where Stand = '" & ti & "'", DB2, 3, 3
            If rs1.State = adStateOpen Then rs1.Close
            rs1.Open "SELECT * FROM Maths2 where Stand = '" & tt & "'", DB2, 3, 3
            
            'rs!perc = Format(((Val(rs1!Area)) - Val(rs!Area)) / Val(rs!Area) * 100, "###,####.00")
           rs!perc = Format(((Val(rs1!Area) - Val(rs!Area)) / Val(rs!Area)) * 100, "###,####.00")
            
            'rs!perc = ((Val(rs1!Area) - Val(rs!Area)) / Val(rs!Area)) * 100
            rs.Update
        End If
      
        
    ti = Val(ti) - 1
   ' End If
    RS3.MoveNext
    Loop
   message ("PROGRAM FOUR")
End Sub

Private Sub txtc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtcontactfriction_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub


Private Sub txtdrivetrainefficiency_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtfrictionalcontactarea_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtmn_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtoilfriction_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtq16_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtQ7_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtspeed_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub


Private Sub txttemperature_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub message(data As String)
Qpn = MsgBox("Calculation of " & data & " - was Succesfully Completed and Reult Generated", vbOKOnly + vbInformation, "CALCULATION MESSAGE")
End Sub
