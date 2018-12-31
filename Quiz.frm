VERSION 5.00
Begin VB.Form Quiz 
   BackColor       =   &H00C0E0FF&
   Caption         =   "QUIZ"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   ForeColor       =   &H80000010&
   LinkTopic       =   "Form1"
   Picture         =   "Quiz.frx":0000
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "SUMMARY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8880
      Width           =   2535
   End
   Begin VB.CommandButton calculator 
      BackColor       =   &H0080C0FF&
      Caption         =   "CALCULATOR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9960
      Width           =   3255
   End
   Begin VB.OptionButton d 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pressure Drop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13920
      TabIndex        =   11
      Top             =   7680
      Width           =   4935
   End
   Begin VB.OptionButton b 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Bolier Efficiency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   10
      Top             =   7080
      Width           =   4935
   End
   Begin VB.OptionButton c 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Temperature Drop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   9
      Top             =   7800
      Width           =   4575
   End
   Begin VB.OptionButton a 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Steam Required"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   8
      Top             =   7080
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "PREV"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Label question 
      BackStyle       =   0  'Transparent
      Caption         =   "Proper sizing of steam pipeline helps in minimising"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   8880
      TabIndex        =   7
      Top             =   3720
      Width           =   11175
   End
   Begin VB.Label qno 
      BackStyle       =   0  'Transparent
      Caption         =   "Question No. 1 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Left:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16200
      TabIndex        =   4
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label ques 
      BackStyle       =   0  'Transparent
      Caption         =   "Question 1 of 20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   1545
      Left            =   6480
      Picture         =   "Quiz.frx":6A17
      Top             =   1320
      Width           =   13560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   20400
      X2              =   5760
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   9480
   End
   Begin VB.Image Image1 
      Height          =   8925
      Left            =   0
      Picture         =   "Quiz.frx":AFD1
      Top             =   0
      Width           =   5535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   20280
      X2              =   5640
      Y1              =   9480
      Y2              =   9480
   End
End
Attribute VB_Name = "Quiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calculator_Click()
On Error GoTo errHandle
Dim a As Double
a = Shell("C:\WINDOWS\System32\calc.exe", vbNormalFocus)
Exit Sub
errHandle:
MsgBox "Unable to run Calculator Utility on your computer", vbInformation, "Error in opening!!!"
Resume Next
End Sub

Private Sub Command5_Click()
Unload Me
summary.Show
End Sub

Private Sub Form_Load()
If course = "MCA" Then
Image2.Picture = LoadPicture("c:\users\nikhil\desktop\my project\images\mca.jpg")
End If
End Sub
