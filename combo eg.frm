VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15885
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   15885
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Apple"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mango"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   765
      Left            =   6480
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10680
      TabIndex        =   1
      Top             =   2280
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6480
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2520
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Add_Click()
'Combo1.AddItem (Text1.Text)
'List1.AddItem (Text1.Text)
'Text1.Text = Check1.Caption
If Check = True Then
Text1.Text = Check1.Caption
End If
End Sub

Private Sub Option1_Click()
If Option1.Enabled = True Then
Text1.Text = Option1.Caption
Else
Text1.Text = Option2.Caption
End If
End Sub
