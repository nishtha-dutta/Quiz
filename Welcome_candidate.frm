VERSION 5.00
Begin VB.Form Welcome_candidate 
   BackColor       =   &H00C0E0FF&
   Caption         =   "WELCOME CANDIDATE"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   3960
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "START QUIZ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8520
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   3375
      Left            =   12960
      Picture         =   "Welcome_candidate.frx":0000
      Top             =   7080
      Width           =   780
   End
   Begin VB.Image Image3 
      Height          =   3315
      Left            =   6360
      Picture         =   "Welcome_candidate.frx":0912
      Top             =   7080
      Width           =   870
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   14160
      X2              =   14160
      Y1              =   6000
      Y2              =   11040
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5880
      X2              =   5880
      Y1              =   6000
      Y2              =   11040
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5880
      X2              =   14160
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Image Image2 
      Height          =   5280
      Left            =   5880
      Picture         =   "Welcome_candidate.frx":130F
      Top             =   6000
      Width           =   8280
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "4. You may quit at any time from the QUIZ."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   4
      Top             =   5520
      Width           =   12375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3. You have 15 minutes to complete the QUIZ and your time starts by clicking on START QUIZ button. "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Top             =   4800
      Width           =   16575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2. There will be no negative marking."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   4080
      Width           =   12375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1. There will be be 20 questions. Each question has four option as answer you have to choose the appropriate one."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   3360
      Width           =   16455
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1560
      X2              =   8520
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rules And Regulations :- "
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1335
      Left            =   1560
      TabIndex        =   0
      Top             =   1920
      Width           =   8895
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   4800
      Picture         =   "Welcome_candidate.frx":4138
      Top             =   120
      Width           =   11190
   End
End
Attribute VB_Name = "Welcome_candidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Quiz.Show
End Sub

Private Sub Command2_Click()
Unload Me
AQS.Show
End Sub

Private Sub Timer1_Timer()
If Image3.Visible = True Then
Image4.Visible = False
Image3.Visible = False
Else
Image4.Visible = True
Image3.Visible = True
End If
End Sub
