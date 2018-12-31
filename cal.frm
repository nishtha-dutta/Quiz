VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      Caption         =   "."
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "/"
      Height          =   495
      Left            =   2520
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "*"
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "-"
      Height          =   495
      Left            =   2520
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command13 
      Caption         =   "+"
      Height          =   2175
      Left            =   1800
      TabIndex        =   12
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "="
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "0"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "c"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim k As Integer
Dim m As Integer
Dim c As String

Private Sub Command1_Click()
Text1.Text = Text1.Text + Command1.Caption
End Sub

Private Sub Command10_Click()
Text1.Text = ""
i = 0
j = 1
x = 0
k = 0
m = 0
End Sub

Private Sub Command11_Click()
Text1.Text = Text1.Text + Command11.Caption
End Sub

Private Sub Command12_Click()
Select Case c
Case "+"
    Text1.Text = Val(Text1.Text) + i
Case "-"
    Text1.Text = k - Val(Text1.Text)
Case "/"
    Text1.Text = m / Val(Text1.Text)
Case "*"
    Text1.Text = Val(Text1.Text) * j
End Select
End Sub

Private Sub Command13_Click()
i = i + Val(Text1.Text)
Text1.Text = ""
c = "+"
j = 1
m = 0
k = 0
End Sub

Private Sub Command14_Click()
k = Val(Text1.Text) - k
Text1.Text = ""
c = "-"
j = 1
i = 0
m = 0
End Sub

Private Sub Command15_Click()
j = j * Val(Text1.Text)
Text1.Text = ""
c = "*"
i = 0
k = 0
m = 0
End Sub

Private Sub Command16_Click()
If (x = 0) Then
m = Val(Text1.Text)
x = x + 1
Else
m = m / Val(Text1.Text)
x = x + 1
End If
Text1.Text = ""
c = "/"
i = 0
j = 1
k = 0
End Sub

Private Sub Command17_Click()
Text1.Text = Text1.Text + Command17.Caption
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text + Command2.Caption
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text + Command3.Caption
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text + Command4.Caption
End Sub

Private Sub Command5_Click()
Text1.Text = Text1.Text + Command5.Caption
End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text + Command6.Caption
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text + Command7.Caption
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text + Command8.Caption
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text + Command9.Caption
End Sub

Private Sub Form_Load()
j = 1
End Sub

