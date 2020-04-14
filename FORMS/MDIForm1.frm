VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FF00FF&
   Caption         =   "THOMPSON VENTURES NIGERIA LIMITED"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10935
      Left            =   0
      ScaleHeight     =   10875
      ScaleWidth      =   11190
      TabIndex        =   0
      Top             =   0
      Width           =   11250
      Begin VB.Timer Timer2 
         Interval        =   10000
         Left            =   14280
         Top             =   1440
      End
      Begin VB.CommandButton Command8 
         Caption         =   "CREATE NEW PARAMETERS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   13
         Top             =   5400
         Width           =   3255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "CLOSE/EXIT "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   12
         Top             =   8880
         Width           =   2895
      End
      Begin VB.CommandButton Command4 
         Caption         =   "REGISTER NEW GOODS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton Command7 
         Caption         =   "GENERATE REPORT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   9
         Top             =   6480
         Width           =   3255
      End
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   120
         Top             =   840
      End
      Begin VB.CommandButton Command6 
         Caption         =   "CALCULATOR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   5
         Top             =   7560
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "MODIFY GOODS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Top             =   4440
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PERFORM SALES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   3
         Top             =   3480
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "RECEIVE INCOMING GOODS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   2
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shop 7 Salma Plaza Lokoja, Kogi State"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9240
         TabIndex        =   15
         Top             =   600
         Width           =   4215
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   14760
         Top             =   1440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1024
         ImageHeight     =   768
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   7
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":240052
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":6000A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":9C00F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":D80148
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":114019A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":13801EC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
         Height          =   1095
         Left            =   240
         TabIndex        =   14
         Top             =   8640
         Width           =   3495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IFY - TOM SUPERMARKET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3720
         TabIndex        =   10
         Top             =   120
         Width           =   14535
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   14040
         TabIndex        =   8
         Top             =   9240
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   4920
         TabIndex        =   7
         Top             =   9240
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"MDIForm1.frx":15C023E
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   10080
         Width           =   41895
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   10
         Height          =   855
         Left            =   0
         Top             =   9960
         Width           =   20175
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   10
         Height          =   8655
         Left            =   120
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   10
         Height          =   1095
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   20175
      End
      Begin VB.Image Image1 
         Height          =   8655
         Left            =   3960
         Picture         =   "MDIForm1.frx":15C02F8
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   16140
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   10
         Height          =   8775
         Left            =   3960
         Top             =   1200
         Width           =   16215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Height          =   11055
         Left            =   120
         TabIndex        =   1
         Top             =   -120
         Width           =   20175
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ICOUNTER As Integer

Private Sub Command1_Click()
FRMRECEIVE.Show vbModal
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.FontSize = 14
Command2.FontSize = 14
Command1.FontSize = 10
Command3.FontSize = 14
Command8.FontSize = 14
Command7.FontSize = 14
Command6.FontSize = 14
End Sub

Private Sub Command2_Click()
FRMSOLD.Show vbModal
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.FontSize = 14
Command2.FontSize = 10
Command1.FontSize = 14
Command3.FontSize = 14
Command8.FontSize = 14
Command7.FontSize = 14
Command6.FontSize = 14
End Sub

Private Sub Command3_Click()
FRMUPDATE.Show vbModal
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.FontSize = 14
Command2.FontSize = 14
Command1.FontSize = 14
Command3.FontSize = 10
Command8.FontSize = 14
Command7.FontSize = 14
Command6.FontSize = 14
End Sub

Private Sub Command4_Click()
FRMNEW.Show vbModal
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.FontSize = 10
Command2.FontSize = 14
Command1.FontSize = 14
Command3.FontSize = 14
Command8.FontSize = 14
Command7.FontSize = 14
Command6.FontSize = 14
End Sub

Private Sub Command5_Click()
Dim J
J = MsgBox("Do you really want to QUIT THIS PROGRAM PLEASE CONFIRM BEFORE..AND BE SURE !!", vbYesNo + vbQuestion, "CONFIRMATION MESSAGE")
If J = vbNo Then
MsgBox "YOU ARE HIGHLY WELCOME BACK ...PLEASE ENJOY THE SYSTEM !!!!!"
Exit Sub
Else
End
End If
End Sub

Private Sub Command6_Click()
Shell ("CALC"), vbMinimizedFocus
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.FontSize = 14
Command2.FontSize = 14
Command1.FontSize = 14
Command3.FontSize = 14
Command8.FontSize = 14
Command7.FontSize = 14
Command6.FontSize = 10
End Sub

Private Sub Command7_Click()
FRMREPORT.Show vbModal
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.FontSize = 14
Command2.FontSize = 14
Command1.FontSize = 14
Command3.FontSize = 14
Command8.FontSize = 14
Command7.FontSize = 10
Command6.FontSize = 14
End Sub

Private Sub Command8_Click()
FRMCREATE.Show vbModal
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.FontSize = 14
Command2.FontSize = 14
Command1.FontSize = 14
Command3.FontSize = 14
Command8.FontSize = 10
Command7.FontSize = 14
Command6.FontSize = 14
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.FontSize = 14
Command2.FontSize = 14
Command1.FontSize = 14
Command3.FontSize = 14
Command8.FontSize = 14
Command7.FontSize = 14
Command6.FontSize = 14
End Sub

Private Sub Timer1_Timer()
If (Label2.Left + Label2.Width) <= 0 Then
     Label2.Left = Me.Width
End If
Label2.Left = Label2.Left - 100
Label4.Caption = Format(Now, "ddd dd mmm, yyyy")
Label3.Caption = Format(Now, "hh:mm:ss: AM/PM")
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

ElseIf ICOUNTER = 7 Then
Image1.Picture = ImageList1.ListImages((CInt(ICOUNTER))).Picture

ICOUNTER = 0
End If

End Sub
