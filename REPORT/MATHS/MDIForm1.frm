VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00404000&
   Caption         =   "ROLLIN STAND CALCULATION SYSTEM"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   150
   ClientWidth     =   9105
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      Height          =   15615
      Left            =   0
      Picture         =   "MDIForm1.frx":0000
      ScaleHeight     =   15555
      ScaleWidth      =   9045
      TabIndex        =   0
      Top             =   0
      Width           =   9105
      Begin VB.Timer Timer2 
         Interval        =   2000
         Left            =   2640
         Top             =   600
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1920
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   492
         ImageHeight     =   329
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":82075
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":90E00
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":94553
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":975CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9AC55
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9EA80
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":BFF45
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":141FCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1682D5
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   1440
         Top             =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "Project Design Present  By:   Name: Isezuo, Lawrence O."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   14520
         TabIndex        =   2
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
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
         Height          =   735
         Left            =   14520
         TabIndex        =   1
         Top             =   5160
         Width           =   5295
      End
      Begin VB.Image Image1 
         Height          =   3930
         Left            =   14520
         Picture         =   "MDIForm1.frx":1ABFFF
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000080&
         Height          =   6015
         Left            =   14160
         TabIndex        =   3
         Top             =   120
         Width           =   5895
      End
   End
   Begin VB.Menu mnucalculate 
      Caption         =   "PERFORM CALCULATION"
   End
   Begin VB.Menu mnuadjustconstants 
      Caption         =   "ADJUST CONSTANTS"
   End
   Begin VB.Menu mnuprintreport 
      Caption         =   "PRINT REPORT"
      Begin VB.Menu rptone 
         Caption         =   "REPORT ONE"
      End
      Begin VB.Menu rpttwo 
         Caption         =   "REPORT TWO"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ICOUNTER As Integer
Private Sub mnuadjustconstants_Click()
Form2.Show
End Sub

Private Sub mnucalculate_Click()
Form1.Show
End Sub

Private Sub rptone_Click()
'display report one
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Maths\maths.mdb")
APPACCESS.DoCmd.OpenReport "maths2", acViewPreview
APPACCESS.Visible = True
End Sub

Private Sub rpttwo_Click()
'display report two
Dim ACCESSAPP As Access.Application
Set APPACCESS = New Access.Application
Set APPACCESS = CreateObject("ACCESS.APPLICATION")
APPACCESS.OpenCurrentDatabase ("C:\Maths\maths.mdb")
APPACCESS.DoCmd.OpenReport "maths2", acViewPreview
APPACCESS.Visible = True
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Format(Now, "ddd dd mmm, yyyy")
Label1.Caption = Label1.Caption & "  -  " & Format(Now, "hh:mm:ss: AM/PM")
End Sub

Private Sub Timer2_Timer()
ICOUNTER = ICOUNTER + 1
    If ICOUNTER = 1 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 2 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 3 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 4 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 5 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 6 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 8 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
    ElseIf ICOUNTER = 9 Then
        Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture
        ICOUNTER = 0
    End If
End Sub
