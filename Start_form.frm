VERSION 5.00
Begin VB.Form Start_form 
   BackColor       =   &H00C0E0FF&
   Caption         =   "WELCOME TO AUTOMATED QUIZ SYSTEM"
   ClientHeight    =   9570
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13020
   FillColor       =   &H80000012&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   13020
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1560
      Top             =   3120
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
      Height          =   1215
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8760
      UseMaskColor    =   -1  'True
      Width           =   3135
   End
   Begin VB.CommandButton Enter 
      BackColor       =   &H0080C0FF&
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Image Image3 
      Height          =   1605
      Left            =   6000
      Picture         =   "Start_form.frx":0000
      Top             =   120
      Width           =   8415
   End
   Begin VB.Image Image2 
      Height          =   2400
      Left            =   1320
      Picture         =   "Start_form.frx":2EF5
      Top             =   1560
      Width           =   18030
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   7440
      Picture         =   "Start_form.frx":970B
      Top             =   4080
      Width           =   5760
   End
End
Attribute VB_Name = "Start_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Enter_Click()
Unload Me
AQS.Show
End Sub

Private Sub Exit_Click()
thank.Show
Unload Me
End Sub


Private Sub Timer1_Timer()
If Image2.Visible = True Then
Image2.Visible = False
Else
Image2.Visible = True
End If
End Sub
