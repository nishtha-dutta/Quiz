VERSION 5.00
Begin VB.Form thank 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7545
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "thank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame thank 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000010&
      Height          =   7275
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.Timer Timer1 
         Interval        =   80
         Left            =   4440
         Top             =   360
      End
      Begin VB.Image Image2 
         Height          =   1425
         Left            =   5040
         Picture         =   "thank.frx":000C
         Top             =   5880
         Width           =   4905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GUIDED BY:-    MR. AMIT TYAGYI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   5880
         Width           =   6255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ALLAHABAD (U.P), MCA (2008-2011), IIMT GR. NOIDA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2400
         TabIndex        =   9
         Top             =   1680
         Width           =   6015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ALLAHABAD (U.P), MCA (2008-2011), IIMT GR. NOIDA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2400
         TabIndex        =   8
         Top             =   5040
         Width           =   6015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ALLAHABAD (U.P), MCA (2008-2011), IIMT GR. NOIDA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   2400
         TabIndex        =   7
         Top             =   3960
         Width           =   6015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ALLAHABAD (U.P), MCA (2008-2011), IIMT GR. NOIDA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   2400
         TabIndex        =   6
         Top             =   2880
         Width           =   6015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " MISS SHAILY TYAGYI,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   3000
         TabIndex        =   5
         Top             =   4560
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MR. ABHISHEK TIWARI, "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   4
         Top             =   3480
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "              MISS PALLAVI SRIVASTAVA, "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   1800
         TabIndex        =   3
         Top             =   2400
         Width           =   6015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MR. NIKHIL DUTTA,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   3480
         TabIndex        =   2
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   1815
         Index           =   3
         Left            =   8040
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   1935
         Index           =   2
         Left            =   360
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   1935
         Index           =   1
         Left            =   8040
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   1815
         Index           =   0
         Left            =   360
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEVELOPERS :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "thank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Timer1_Timer()
If Image2.Visible = True Then
Image2.Visible = False
Else
Image2.Visible = True
End If
If a = 20 Then
Unload Me
Else
a = a + 1
End If
End Sub
