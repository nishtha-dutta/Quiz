VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Sch_mgmt 
   BackColor       =   &H00C0E0FF&
   Caption         =   "SCHEDULE MANAGEMENT"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   20250
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Back"
      Top             =   6480
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Search record"
      Top             =   5400
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "First Record"
      Top             =   8520
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
      TabIndex        =   14
      ToolTipText     =   "Previous Record"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
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
      TabIndex        =   13
      ToolTipText     =   "Update changes"
      Top             =   9240
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
      TabIndex        =   12
      ToolTipText     =   "Delete record"
      Top             =   9240
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
      TabIndex        =   11
      ToolTipText     =   "Next Record"
      Top             =   8520
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
      TabIndex        =   10
      ToolTipText     =   "Last Record"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1935
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
      TabIndex        =   9
      ToolTipText     =   "Back"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1935
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
      Top             =   9240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox time 
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
      Left            =   10080
      TabIndex        =   3
      Top             =   6120
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
      ItemData        =   "Sch_mgmt.frx":0000
      Left            =   10080
      List            =   "Sch_mgmt.frx":0025
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "---Select one---"
      Top             =   3120
      Width           =   4095
   End
   Begin VB.TextBox room 
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
      Left            =   10080
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker dob 
      Height          =   495
      Left            =   10080
      TabIndex        =   0
      Top             =   4080
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
      Format          =   86900737
      CurrentDate     =   40253
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
      Left            =   3600
      TabIndex        =   17
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   3720
      Picture         =   "Sch_mgmt.frx":0088
      Top             =   360
      Width           =   14220
   End
   Begin VB.Label date_label 
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
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label room_label 
      BackStyle       =   0  'Transparent
      Caption         =   "ROOM"
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
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   2895
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
      Left            =   6120
      TabIndex        =   5
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label time_label 
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
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   6240
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "Sch_mgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub Form_Load()
Set db = OpenDatabase("c:\users\nikhil\desktop\my project\quiz.mdb")
Set rs = db.OpenRecordset("schedule")
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
If rs.Fields(2) = branch.Text Then
time.Text = rs.Fields(1)
room.Text = rs.Fields(3)
dob.Value = rs.Fields(0)
branch.Text = rs.Fields(2)
End If
End If
End Sub

Private Sub search_Click()
rs.MoveFirst
If rs.EOF Then
MsgBox "Table Is Empty...!", vbCritical
Exit Sub
Else
    If branch.Text = "---Select one---" Then
    MsgBox "Please select desired branch !! ", vbCritical
    Exit Sub
    Else
    rs.MoveFirst
    While Not rs.EOF
        If rs.Fields(2) = branch.Text And (rs.Fields(0) - VBA.Date) >= 0 Then
        time.Text = rs.Fields(1)
        dob.Value = rs.Fields(0)
        room.Text = rs.Fields(3)
            Label1.Visible = False
            room_label.Visible = True
            date_label.Visible = True
            time_label.Visible = True
            dob.Visible = True
            time.Visible = True
            room.Visible = True
            Prev.Visible = True
            back.Visible = True
            add.Visible = True
            next_button.Visible = True
            Delete.Visible = True
            Update.Visible = True
            First.Visible = True
            Last.Visible = True
            search.Visible = False
            back1.Visible = False
        Exit Sub
        End If
        rs.MoveNext
    Wend
    If rs.EOF Then
    MsgBox "Record Not Found...!", vbCritical
    End If
    End If
End If
End Sub

Private Sub back_Click()
Label1.Visible = True
room_label.Visible = False
date_label.Visible = False
time_label.Visible = False
dob.Visible = False
time.Visible = False
room.Visible = False
Prev.Visible = False
back.Visible = False
add.Visible = False
next_button.Visible = False
Delete.Visible = False
Update.Visible = False
First.Visible = False
Last.Visible = False
search.Visible = True
back1.Visible = True
End Sub

Private Sub back1_Click()
Unload Me
Welcome_admin.Show
End Sub


