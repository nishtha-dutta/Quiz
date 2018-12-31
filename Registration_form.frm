VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Registration_form 
   BackColor       =   &H00C0E0FF&
   Caption         =   "REGISTRATION FORM"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker dob 
      Height          =   495
      Left            =   10080
      TabIndex        =   17
      Top             =   5400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   36896769
      CurrentDate     =   40253
   End
   Begin VB.OptionButton female 
      BackColor       =   &H00C0E0FF&
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11640
      TabIndex        =   16
      Top             =   6360
      Width           =   1575
   End
   Begin VB.OptionButton male 
      BackColor       =   &H00C0E0FF&
      Caption         =   "MALE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10080
      TabIndex        =   15
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox email 
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
      Left            =   10080
      TabIndex        =   14
      Top             =   7800
      Width           =   4215
   End
   Begin VB.TextBox contact 
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
      Left            =   10080
      TabIndex        =   13
      Top             =   6960
      Width           =   4215
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
      ItemData        =   "Registration_form.frx":0000
      Left            =   10080
      List            =   "Registration_form.frx":0025
      Sorted          =   -1  'True
      TabIndex        =   8
      Text            =   "---Select one---"
      Top             =   4440
      Width           =   4215
   End
   Begin VB.TextBox Roll_no 
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
      Left            =   10080
      TabIndex        =   7
      Top             =   3600
      Width           =   4215
   End
   Begin VB.TextBox cname 
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
      Left            =   10080
      TabIndex        =   6
      Top             =   2760
      Width           =   4215
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H0080C0FF&
      Caption         =   "BACK"
      Height          =   735
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton Clear 
      BackColor       =   &H0080C0FF&
      Caption         =   "CLEAR"
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Clear All"
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Submit 
      BackColor       =   &H0080C0FF&
      Caption         =   "SUBMIT"
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Submit Details"
      Top             =   9600
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5760
      TabIndex        =   12
      Top             =   6240
      Width           =   3975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "E-MAIL ADDRESS"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5760
      TabIndex        =   11
      Top             =   7920
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT NO."
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5760
      TabIndex        =   10
      Top             =   7080
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5760
      TabIndex        =   9
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   2040
      Left            =   3600
      Picture         =   "Registration_form.frx":0088
      Top             =   480
      Width           =   13815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ROLL NO."
      Height          =   855
      Left            =   5760
      TabIndex        =   2
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "COURSE/BRANCH"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5760
      TabIndex        =   1
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      Height          =   855
      Left            =   5760
      TabIndex        =   0
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "Registration_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub Clear_Click()
male.Value = False
female.Value = False
Roll_no.Text = ""
cname.Text = ""
contact.Text = ""
email.Text = ""
dob.Value = "16 - 3 - 2010"
branch.Text = "---Select one---"
End Sub

Private Sub Exit_Click()
AQS.Show
Unload Me
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("c:\users\nikhil\desktop\my project\quiz.mdb")
Set rs = db.OpenRecordset("candidate")
End Sub

Private Sub Submit_Click()
If contact.Text = "" Or email.Text = "" Or Roll_no.Text = "" Or cname.Text = "" Or branch.Text = "---Select one---" Or (male.Value = False And female.Value = False) Then
MsgBox "Please Fill All Enteries !!", vbCritical
Exit Sub
End If
rs.AddNew
rs.Fields(6) = Roll_no.Text
rs.Fields(0) = cname.Text
rs.Fields(1) = branch.Text
rs.Fields(3) = dob.Value
rs.Fields(4) = contact.Text
rs.Fields(5) = email.Text
If male.Value = True Then
rs.Fields(2) = "M"
Else
rs.Fields(2) = "F"
End If
rs.Update
MsgBox "Registration Successfull !!", vbInformation
End Sub
