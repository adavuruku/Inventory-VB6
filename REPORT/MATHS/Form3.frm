VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF8080&
   Caption         =   "MAIN WINDOW"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form3"
   ScaleHeight     =   3435
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "PRINT RESULT"
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
      Left            =   3600
      MaskColor       =   &H000080FF&
      TabIndex        =   2
      ToolTipText     =   "Click here to Adjust Constatnts"
      Top             =   2040
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADJUST CONSTANTS"
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
      Left            =   7200
      MaskColor       =   &H000080FF&
      TabIndex        =   1
      ToolTipText     =   "Click here to Adjust Constatnts"
      Top             =   720
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      Caption         =   "PERFORM CALCULATIONS"
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
      Left            =   120
      MaskColor       =   &H000080FF&
      TabIndex        =   0
      ToolTipText     =   "Click here to perform calculations"
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub
