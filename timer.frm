VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   240
      Top             =   720
   End
   Begin VB.Label Label1 
      Caption         =   "NIKHIL DUTTA"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Label1.Visible = True Then
Label1.Visible = False
Else
Label1.Visible = True
End If
End Sub
