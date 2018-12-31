VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Candidate_info_mgmt 
   BackColor       =   &H00C0E0FF&
   Caption         =   "CANDIDATE INFORMATION MANAGEMENT"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cancel 
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Next Record"
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton ok 
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Next Record"
      Top             =   5160
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dob 
      Height          =   495
      Left            =   10800
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
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
      Format          =   89260033
      CurrentDate     =   40253
   End
   Begin VB.TextBox rollno 
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
      Left            =   10800
      TabIndex        =   20
      Top             =   3840
      Width           =   4095
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
      Left            =   13200
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
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
      Left            =   10920
      TabIndex        =   18
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox email 
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
      Left            =   10800
      TabIndex        =   17
      Top             =   7440
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox contact 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10800
      TabIndex        =   16
      Top             =   6720
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
      ItemData        =   "Candidate_info_mgmt.frx":0000
      Left            =   10800
      List            =   "Candidate_info_mgmt.frx":0025
      Sorted          =   -1  'True
      TabIndex        =   11
      Text            =   "---Select one---"
      Top             =   4560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox cname 
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
      Left            =   10800
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   4095
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Back"
      Top             =   9600
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Last Record"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Next_button 
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Next Record"
      Top             =   8880
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete record"
      Top             =   9600
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Update changes"
      Top             =   9600
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Previous Record"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "First Record"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select  :-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   23
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label email_label 
      BackStyle       =   0  'Transparent
      Caption         =   "E-MAIL ADDRESS"
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
      Left            =   6720
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label contact_label 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT NO."
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
      Left            =   6720
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label gender_label 
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
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
      Left            =   6720
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label dob_label 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
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
      Left            =   6720
      TabIndex        =   12
      Top             =   6120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1770
      Left            =   2400
      Picture         =   "Candidate_info_mgmt.frx":0088
      Top             =   240
      Width           =   16215
   End
   Begin VB.Label branch_label 
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
      Left            =   6720
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label cname_label 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
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
      Left            =   6720
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ROLL NUMBER"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   3960
      Width           =   3375
   End
End
Attribute VB_Name = "Candidate_info_mgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub back_Click()
rollno.Text = ""
cname_label.Visible = False
dob_label.Visible = False
gender_label.Visible = False
email_label.Visible = False
contact_label.Visible = False
branch_label.Visible = False
cname.Visible = False
branch.Visible = False
male.Visible = False
female.Visible = False
dob.Visible = False
email.Visible = False
contact.Visible = False
next_button.Visible = False
Prev.Visible = False
First.Visible = False
Last.Visible = False
Update.Visible = False
Delete.Visible = False
back.Visible = False
ok.Visible = True
cancel.Visible = True
Label2.Visible = True
End Sub

Private Sub cancel_Click()
Unload Me
Welcome_admin.Show
End Sub

Private Sub Delete_Click()
p = MsgBox("Are you sure to delete the record ??", vbQuestion + vbYesNo)
If p = vbYes Then
rs.Delete
MsgBox "Record deleted successfully !!", vbInformation
End If
End Sub

Private Sub First_Click()
rs.MoveFirst
rollno.Text = rs.Fields(6)
cname.Text = rs.Fields(0)
branch.Text = rs.Fields(1)
dob.Value = rs.Fields(3)
contact.Text = rs.Fields(4)
email.Text = rs.Fields(5)
If rs.Fields(2) = "M" Then
male.Value = True
Else
female.Value = True
End If
End Sub

Private Sub Last_Click()
rs.MoveLast
rollno.Text = rs.Fields(6)
cname.Text = rs.Fields(0)
branch.Text = rs.Fields(1)
dob.Value = rs.Fields(3)
contact.Text = rs.Fields(4)
email.Text = rs.Fields(5)
If rs.Fields(2) = "M" Then
male.Value = True
Else
female.Value = True
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("c:\users\nikhil\desktop\my project\quiz.mdb")
Set rs = db.OpenRecordset("candidate")

End Sub

Private Sub next_button_Click()
If rs.EOF = True Then
rs.MoveLast
l:
MsgBox "No more records !!", vbCritical
Else
rs.MoveNext
If rs.EOF = True Then
GoTo l
End If
rollno.Text = rs.Fields(6)
cname.Text = rs.Fields(0)
branch.Text = rs.Fields(1)
dob.Value = rs.Fields(3)
contact.Text = rs.Fields(4)
email.Text = rs.Fields(5)
If rs.Fields(2) = "M" Then
male.Value = True
Else
female.Value = True
End If
End If
End Sub


Private Sub ok_Click()
If rs.RecordCount = 0 Then
    MsgBox "Table is EMPTY !!", vbCritical
    Exit Sub
ElseIf rollno.Text = "" Then
    MsgBox "Please enter the ROLL NO. !!", vbCritical
    Exit Sub
Else
rs.MoveFirst
    While Not rs.EOF
        If rs.Fields(6) = rollno.Text Then

            cname_label.Visible = True
            dob_label.Visible = True
            gender_label.Visible = True
            email_label.Visible = True
            contact_label.Visible = True
            branch_label.Visible = True
            cname.Visible = True
            branch.Visible = True
            male.Visible = True
            female.Visible = True
            dob.Visible = True
            email.Visible = True
            contact.Visible = True
            next_button.Visible = True
            Prev.Visible = True
            First.Visible = True
            Last.Visible = True
            Update.Visible = True
            Delete.Visible = True
            back.Visible = True
            ok.Visible = False
            cancel.Visible = False
            Label2.Visible = False

            rollno.Text = rs.Fields(6)
            cname.Text = rs.Fields(0)
            branch.Text = rs.Fields(1)
            dob.Value = rs.Fields(3)
            contact.Text = rs.Fields(4)
            email.Text = rs.Fields(5)
            If rs.Fields(2) = "M" Then
                male.Value = True
            Else
                female.Value = True
            End If
            Exit Sub
        End If
        rs.MoveNext
    Wend
    If rs.EOF Then
    MsgBox "No such record FOUND !!", vbCritical
    End If
End If
End Sub

Private Sub prev_Click()
If rs.BOF = True Then
rs.MoveFirst
l:
MsgBox "No more records !!", vbCritical
Else
rs.MovePrevious
If rs.BOF = True Then
GoTo l
End If
rollno.Text = rs.Fields(6)
cname.Text = rs.Fields(0)
branch.Text = rs.Fields(1)
dob.Value = rs.Fields(3)
contact.Text = rs.Fields(4)
email.Text = rs.Fields(5)
If rs.Fields(2) = "M" Then
male.Value = True
Else
female.Value = True
End If
End If
End Sub

Private Sub Update_Click()
If contact.Text = "" Or email.Text = "" Or rollno.Text = "" Or cname.Text = "" Then
MsgBox "Please Fill All Enteries !!", vbCritical
Exit Sub
End If
rs.Edit
rs.Fields(6) = rollno.Text
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
p = MsgBox("Are you sure to update the record ??", vbQuestion + vbYesNo)
If p = vbYes Then
rs.Update
MsgBox "Changes saved successfully !!", vbInformation
Else
Exit Sub
End If
End Sub
