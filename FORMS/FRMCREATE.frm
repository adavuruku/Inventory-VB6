VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMCREATE 
   BackColor       =   &H00004000&
   Caption         =   "ADD NEW PARAMETERS FORM"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5880
      Top             =   8520
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
   Begin VB.Frame Frame3 
      BackColor       =   &H000080FF&
      Caption         =   "CREATE NEW MEASUREMENT UNIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   11
      Top             =   5880
      Width           =   6975
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   5895
      End
      Begin VB.CommandButton Command6 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   8
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   9
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Height          =   975
         Left            =   480
         TabIndex        =   14
         Top             =   1440
         Width           =   5895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "CREATE NEW GOODS TYPE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   6975
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   5895
      End
      Begin VB.CommandButton Command4 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   5
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   6
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   975
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   5895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "CREATE NEW ITEM LOCATION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   2
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   6015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Height          =   975
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   6015
      End
   End
End
Attribute VB_Name = "FRMCREATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please type in the new Location you want to create here.. before saving it", vbInformation, "WRONG INFORMATION GIVEN"
Text1.SetFocus
Exit Sub
End If

Dim m
m = MsgBox("Do you really want to SAVE THIS  " & Text1.Text & "  AS A NEW LOCATION", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If m = vbNo Then
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Exit Sub

Else
Dim DB20 As New ADODB.CONNECTION
Dim RS40 As New ADODB.Recordset
'Dim RS2 As New ADODB.Recordset
If DB20.State = adStateOpen Then DB20.Close
DB20.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS40.Open "SELECT * FROM LOCATION Where Location ='" & Trim(Text1.Text) & "'", DB20, adOpenDynamic, adLockOptimistic
If Not RS40.EOF Then
MsgBox "These Location " & Text1.Text & " You want to create is already existing ...please is not recommended to create two same group.. please either add one or more phrase to differentiate it", vbInformation, "EXISTING RECORDS DETECTED"
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Exit Sub

Else
Dim DB As New ADODB.CONNECTION
Dim RS4 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS2.Open "SELECT * FROM LOCATION", DB, adOpenDynamic, adLockOptimistic
RS2.AddNew
RS2!Location = UCase(Text1.Text)

RS2!DATECREATE = Format(Now, "DD/MM/YYYY")
RS2.Update
MsgBox "New Location has been created succesfully", vbInformation, "New Location created"
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End If
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Text2.Text = ""
Text2.SetFocus
End Sub

Private Sub Command4_Click()
If Text2.Text = "" Then
MsgBox "Please type in the new GROUP OF GOODS you want to create here.. before saving it", vbInformation, "WRONG INFORMATION GIVEN"
Text2.SetFocus
Exit Sub
End If

Dim m
m = MsgBox("Do you really want to SAVE  " & Text2.Text & "   AS A NEW GROUP OF GOODS", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If m = vbNo Then
Text2.SetFocus
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
Exit Sub
Else

Dim DB25 As New ADODB.CONNECTION
Dim RS45 As New ADODB.Recordset
'Dim RS2 As New ADODB.Recordset
If DB25.State = adStateOpen Then DB25.Close
DB25.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS45.Open "SELECT * FROM GoodsType Where Type ='" & Trim(Text2.Text) & "'", DB25, adOpenDynamic, adLockOptimistic
If Not RS45.EOF Then
MsgBox "These Goods Type " & Text2.Text & " You want to create is already existing ...please is not recommended to create TWO SAME GROUP.. Please either add one or more phrase to differentiate it", vbInformation, "EXISTING RECORDS DETECTED"
Text2.SetFocus
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
Exit Sub

Else
Dim DB1 As New ADODB.CONNECTION
Dim RS44 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
DB1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS44.Open "SELECT * FROM GoodsType", DB1, adOpenDynamic, adLockOptimistic
RS44.AddNew
RS44!Type = UCase(Trim(Text2.Text))
RS44!DATECREATE = Format(Now, "DD/MM/YYYY")
RS44!GOODSCODE = "1001"
RS44!REP = Trim(Text2.Text)
RS44.Update
MsgBox "New GOODS TYPE has been created succesfully", vbInformation, "New GOODS TYPE created"
Text2.SetFocus
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End If
End If
End Sub

Private Sub Command5_Click()
Text3.Text = ""
Text3.SetFocus
End Sub

Private Sub Command6_Click()
If Text3.Text = "" Then
MsgBox "Please type in the new MEASUREMENT PARAMETERS you want to create here.. before saving it", vbInformation, "WRONG INFORMATION GIVEN"
Text3.SetFocus
Exit Sub
End If

Dim m
m = MsgBox("Do you really want to SAVE  " & Text3.Text & "   AS A NEW MEASUREMENT PARAMETERS", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If m = vbNo Then
Text3.SetFocus
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
Exit Sub
Else

Dim DB250 As New ADODB.CONNECTION
Dim RS450 As New ADODB.Recordset
'Dim RS2 As New ADODB.Recordset
If DB250.State = adStateOpen Then DB250.Close
DB250.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS450.Open "SELECT * FROM Unit Where measures ='" & Trim(Text3.Text) & "'", DB250, adOpenDynamic, adLockOptimistic
If Not RS450.EOF Then
MsgBox "These MEASUREMENT PARAMETERS, " & Text3.Text & " You want to create is already existing ...please is not recommended to create TWO SAME MEASUREMENT PARAMETERS.. please either add one or more phrase to differentiate it", vbInformation, "EXISTING RECORDS DETECTED"
Text3.SetFocus
Text3.SelStart = 0
Text3.SelLength = Len(Text2.Text)
Exit Sub

Else
Dim DB109 As New ADODB.CONNECTION
Dim RS109 As New ADODB.Recordset
'Dim RS2 As New ADODB.Recordset
DB109.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Market\DATABASE\market.mdb;Persist Security Info=False"
RS109.Open "SELECT * FROM UNIT", DB109, adOpenDynamic, adLockOptimistic
RS109.AddNew
RS109!measures = UCase(Trim(Text3.Text))
RS109!DATECREATE = Format(Now, "DD/MM/YYYY")
'RS4!GOODSCODE = "1001"
'RS4!REP = Trim(Text2.Text)
RS109.Update
MsgBox "New MEASUREMENT PARAMETERS has been created succesfully", vbInformation, "New MEASUREMENT PARAMETERS created"
Text3.SetFocus
Text3.SelStart = 0
Text3.SelLength = Len(Text2.Text)
End If
End If
End Sub

Private Sub Form_Load()
Me.Top = 1200
Me.Left = 5000
End Sub
