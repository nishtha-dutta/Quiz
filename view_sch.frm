VERSION 5.00
Begin VB.Form view_sch 
   BackColor       =   &H00C0E0FF&
   Caption         =   "VIEW SCHEDULE"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   855
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   3255
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
      ItemData        =   "view_sch.frx":0000
      Left            =   12360
      List            =   "view_sch.frx":0025
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "---Select one---"
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   8160
      TabIndex        =   31
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label room 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   13320
      TabIndex        =   30
      Top             =   8040
      Width           =   3255
   End
   Begin VB.Label room 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   13320
      TabIndex        =   29
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label room 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   13320
      TabIndex        =   28
      Top             =   6840
      Width           =   3255
   End
   Begin VB.Label room 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   13320
      TabIndex        =   27
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label room 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   13320
      TabIndex        =   26
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label room 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   13320
      TabIndex        =   25
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label room 
      BackStyle       =   0  'Transparent
      Caption         =   "ROOM NO."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   13320
      TabIndex        =   24
      Top             =   4320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   9840
      TabIndex        =   23
      Top             =   8040
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   9840
      TabIndex        =   22
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   9840
      TabIndex        =   21
      Top             =   6840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   9840
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   9840
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9840
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9840
      TabIndex        =   17
      Top             =   4320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   6840
      TabIndex        =   16
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6840
      TabIndex        =   15
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   6840
      TabIndex        =   14
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   6840
      TabIndex        =   13
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   6840
      TabIndex        =   12
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6840
      TabIndex        =   11
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6840
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label sno 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   4800
      TabIndex        =   9
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label sno 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4800
      TabIndex        =   8
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label sno 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4800
      TabIndex        =   7
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label sno 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   4800
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label sno 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   5
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label sno 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   4
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label sno 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "S No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   16680
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   16680
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   16680
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   16680
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   16680
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   16680
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   16680
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   16680
      X2              =   16680
      Y1              =   4200
      Y2              =   8520
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   13200
      X2              =   13200
      Y1              =   4200
      Y2              =   8520
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   4680
      Y1              =   4200
      Y2              =   8520
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   9720
      X2              =   9720
      Y1              =   4200
      Y2              =   8520
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   6720
      X2              =   6720
      Y1              =   4200
      Y2              =   8520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   16680
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select :-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5040
      TabIndex        =   1
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   2790
      Left            =   5760
      Picture         =   "view_sch.frx":0088
      Top             =   120
      Width           =   9435
   End
End
Attribute VB_Name = "view_sch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
Unload Me
AQS.Show
End Sub

Private Sub branch_Click()
room(0).Visible = True
sno(0).Visible = True
date(0).Visible = True
time(0).Visible = True
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Line5.Visible = True
Line6.Visible = True
Line7.Visible = True
Line8.Visible = True
Line9.Visible = True
Line10.Visible = True
Line11.Visible = True
Line12.Visible = True
Line13.Visible = True
End Sub

