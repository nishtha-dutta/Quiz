VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton back1 
      BackColor       =   &H0080C0FF&
      Caption         =   "CANCEL"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Search record"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton search1 
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
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Search record"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter Your :-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back1_Click()
back1.Visible = True
search.Visible = True
cname_label.Visible = False
cname.Visible = False
branch_label.Visible = False
branch.Visible = False
contact_label.Visible = False
contact.Visible = False
dob.Visible = False
dob_label.Visible = False
email.Visible = False
email_label.Visible = False
gender_label.Visible = False
male.Visible = False
female.Visible = False
First.Visible = False
Last.Visible = False
Prev.Visible = False
Delete.Visible = False
Update.Visible = False
back.Visible = False
Label8.Visible = True
Next_button.Visible = False
End Sub


Private Sub search1_Click()
back1.Visible = True
search1.Visible = True
search.Visible = False
Label8.Visible = False
cname_label.Visible = True
cname.Visible = True
branch_label.Visible = True
branch.Visible = True
contact_label.Visible = True
contact.Visible = True
dob.Visible = True
dob_label.Visible = True
email.Visible = True
email_label.Visible = True
gender_label.Visible = True
male.Visible = True
female.Visible = True
First.Visible = True
Last.Visible = True
Prev.Visible = True
Delete.Visible = True
Update.Visible = True
back.Visible = True
Next_button.Visible = True
End Sub
