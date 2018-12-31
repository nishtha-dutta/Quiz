VERSION 5.00
Begin VB.Form Welcome_admin 
   BackColor       =   &H00C0E0FF&
   Caption         =   "WELCOME ADMINISTRATOR"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9240
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "REPORT GENERATION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   10935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "QUIZ DATABASE MANAGEMENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   10935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "SCHEDULE MANAGEMENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   10935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "CANDIDATE INFORMATION MANAGEMENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   10935
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   3960
      Picture         =   "Welcome_admin.frx":0000
      Top             =   720
      Width           =   13830
   End
End
Attribute VB_Name = "Welcome_admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Candidate_info_mgmt.Show
End Sub

Private Sub Command2_Click()
Unload Me
Sch_mgmt.Show
End Sub

Private Sub Command3_Click()
Unload Me
Quiz_mgmt.Show
End Sub

Private Sub Command5_Click()
Unload Me
AQS.Show
End Sub

