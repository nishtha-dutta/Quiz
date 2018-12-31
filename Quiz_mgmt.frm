VERSION 5.00
Begin VB.Form Quiz_mgmt 
   BackColor       =   &H00C0E0FF&
   Caption         =   "QUIZ MANAGEMENT"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton back1 
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
      Height          =   735
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Exit"
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox qno 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10320
      TabIndex        =   31
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox correct 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   12240
      TabIndex        =   27
      Top             =   7080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox correct 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   5520
      TabIndex        =   26
      Top             =   7080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox correct 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   12240
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox correct 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   5520
      TabIndex        =   24
      Top             =   6360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox option_text 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   12240
      TabIndex        =   23
      Top             =   5640
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox option_text 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   5520
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox ques 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5520
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   10815
   End
   Begin VB.TextBox option_text 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   5520
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ComboBox branch 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Quiz_mgmt.frx":0000
      Left            =   10320
      List            =   "Quiz_mgmt.frx":0025
      Sorted          =   -1  'True
      TabIndex        =   10
      Text            =   "---Select one---"
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox option_text 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   12240
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton add 
      BackColor       =   &H0080C0FF&
      Caption         =   "ADD"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Delete record"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton back 
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
      Height          =   735
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Last 
      BackColor       =   &H0080C0FF&
      Caption         =   "LAST"
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
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Last Record"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton next_button 
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Next Record"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Delete 
      BackColor       =   &H0080C0FF&
      Caption         =   "DELETE"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Delete record"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Update 
      BackColor       =   &H0080C0FF&
      Caption         =   "UPADTE"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Update changes"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Prev 
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Previous Record"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton First 
      BackColor       =   &H0080C0FF&
      Caption         =   "FIRST"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "First Record"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton search 
      BackColor       =   &H0080C0FF&
      Caption         =   "SEARCH"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Search record"
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label correct_label 
      BackStyle       =   0  'Transparent
      Caption         =   "CORRECT 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   3
      Left            =   3000
      TabIndex        =   30
      Top             =   6360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label correct_label 
      BackStyle       =   0  'Transparent
      Caption         =   "CORRECT 4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   2
      Left            =   9720
      TabIndex        =   29
      Top             =   7080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label correct_label 
      BackStyle       =   0  'Transparent
      Caption         =   "CORRECT 3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   1
      Left            =   3000
      TabIndex        =   28
      Top             =   7080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label correct_label 
      BackStyle       =   0  'Transparent
      Caption         =   "CORRECT 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   9720
      TabIndex        =   20
      Top             =   6360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label option_label 
      BackStyle       =   0  'Transparent
      Caption         =   "OPTION 4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   3
      Left            =   9720
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label option_label 
      BackStyle       =   0  'Transparent
      Caption         =   "OPTION 3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   2
      Left            =   3000
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label option_label 
      BackStyle       =   0  'Transparent
      Caption         =   "OPTION 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   1
      Left            =   9720
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   5160
      Picture         =   "Quiz_mgmt.frx":0088
      Top             =   120
      Width           =   11505
   End
   Begin VB.Label option_label 
      BackStyle       =   0  'Transparent
      Caption         =   "OPTION 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   3000
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "COURSE/BRANCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   15
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label ques_label 
      BackStyle       =   0  'Transparent
      Caption         =   "QUESTION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label qno_label 
      BackStyle       =   0  'Transparent
      Caption         =   "QUES NO."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select :-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   3000
      TabIndex        =   12
      Top             =   1920
      Width           =   3975
   End
End
Attribute VB_Name = "Quiz_mgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
Label1.Visible = True
back1.Visible = True
qno_label.Visible = False
ques_label.Visible = False
option_label(3).Visible = False
option_label(0).Visible = False
option_label(1).Visible = False
option_label(2).Visible = False
correct_label(3).Visible = False
correct_label(0).Visible = False
correct_label(1).Visible = False
correct_label(2).Visible = False
qno.Visible = False
ques.Visible = False
correct(3).Visible = False
correct(0).Visible = False
correct(1).Visible = False
correct(2).Visible = False
option_text(3).Visible = False
option_text(0).Visible = False
option_text(1).Visible = False
option_text(2).Visible = False
Prev.Visible = False
back.Visible = False
add.Visible = False
next_button.Visible = False
Delete.Visible = False
Update.Visible = False
First.Visible = False
Last.Visible = False
search.Visible = True
End Sub

Private Sub back1_Click()
Unload Me
Welcome_admin.Show
End Sub

Private Sub search_Click()
search.Visible = False
Label1.Visible = False
back1.Visible = False
qno_label.Visible = True
ques_label.Visible = True
option_label(3).Visible = True
option_label(0).Visible = True
option_label(1).Visible = True
option_label(2).Visible = True
correct_label(3).Visible = True
correct_label(0).Visible = True
correct_label(1).Visible = True
correct_label(2).Visible = True
qno.Visible = True
ques.Visible = True
correct(3).Visible = True
correct(0).Visible = True
correct(1).Visible = True
correct(2).Visible = True
option_text(3).Visible = True
option_text(0).Visible = True
option_text(1).Visible = True
option_text(2).Visible = True
Prev.Visible = True
back.Visible = True
add.Visible = True
next_button.Visible = True
Delete.Visible = True
Update.Visible = True
First.Visible = True
Last.Visible = True
End Sub

