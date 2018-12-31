VERSION 5.00
Begin VB.Form AQS 
   BackColor       =   &H00C0E0FF&
   Caption         =   "AUTOMATED QUIZ SYSTEM"
   ClientHeight    =   3060
   ClientLeft      =   195
   ClientTop       =   495
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox combo_type 
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
      ItemData        =   "AQS.frx":0000
      Left            =   10320
      List            =   "AQS.frx":000A
      TabIndex        =   12
      Text            =   "ADMIN"
      ToolTipText     =   "Select your type"
      Top             =   8280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Register 
      BackColor       =   &H0080C0FF&
      Caption         =   "REGISTER"
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Candidate Registration"
      Top             =   9360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Submit 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Submit"
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Pwd 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   10320
      PasswordChar    =   "*"
      TabIndex        =   5
      ToolTipText     =   "Not more than 10 chars"
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox id 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   4
      ToolTipText     =   "User name"
      Top             =   6720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H0080C0FF&
      Caption         =   "EXIT"
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
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   9360
      Width           =   2295
   End
   Begin VB.CommandButton Q_login 
      BackColor       =   &H0080C0FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Login"
      Top             =   4320
      Width           =   8895
   End
   Begin VB.CommandButton Q_result 
      BackColor       =   &H0080C0FF&
      Caption         =   "QUIZ RESULT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "View Result"
      Top             =   3120
      Width           =   8895
   End
   Begin VB.CommandButton Q_schedule 
      BackColor       =   &H0080C0FF&
      Caption         =   "QUIZ SCHEDULE"
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
      Left            =   5520
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "View Schedule"
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   13680
      TabIndex        =   13
      ToolTipText     =   "Candidate Registration"
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   840
      Picture         =   "AQS.frx":0020
      Top             =   3000
      Width           =   4755
   End
   Begin VB.Label New_user 
      BackStyle       =   0  'Transparent
      Caption         =   "NEW USER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   13800
      TabIndex        =   11
      ToolTipText     =   "Candidate Registration"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label utype 
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE"
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
      TabIndex        =   8
      Top             =   8400
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label upwd 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
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
      TabIndex        =   7
      Top             =   7560
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label uid 
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME"
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
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "AQS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset


Private Sub Exit_Click()
Unload Me
thank.Show
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("c:\users\nikhil\desktop\my project\quiz.mdb")
Set rs = db.OpenRecordset("User")
Set rs1 = db.OpenRecordset("candidate")
End Sub

Private Sub Label1_Click()
input_form.Show
End Sub

Private Sub New_user_Click()
Unload Me
Registration_form.Show
End Sub

Private Sub Q_login_Click()
Label1.Visible = True
uid.Visible = True
Pwd.Visible = True
id.Visible = True
upwd.Visible = True
Submit.Visible = True
Register.Visible = True
New_user.Visible = True
combo_type.Visible = True
utype.Visible = True
End Sub

Private Sub Q_result_Click()
Unload Me
Result.Show
End Sub

Private Sub Q_schedule_Click()
Unload Me
view_sch.Show
End Sub

Private Sub Register_Click()
Unload Me
Registration_form.Show
End Sub

Private Sub Submit_Click()
' while loop should be added
If Pwd.Text = rs1.Fields(6) And id.Text = rs1.Fields(0) And combo_type.Text = "CANDIDATE" Then
course = rs1.Fields(1)
rs.AddNew
rs.Fields(0) = id.Text
rs.Fields(1) = Pwd.Text
rs.Fields(2) = Time
rs.Update
Unload Me
Welcome_candidate.Show
ElseIf Pwd.Text = rs.Fields(1) And id.Text = rs.Fields(0) And combo_type.Text = "ADMIN" Then
Unload Me
Welcome_admin.Show
Else
MsgBox ("Please check your user name and password")
End If
End Sub
