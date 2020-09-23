VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   8280
      Top             =   3840
   End
   Begin VB.Timer Timer5 
      Interval        =   200
      Left            =   5280
      Top             =   2760
   End
   Begin VB.Timer Timer4 
      Interval        =   20
      Left            =   5280
      Top             =   2160
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   5280
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   5280
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5280
      Top             =   120
   End
   Begin Project1.simpleprogressbar simpleprogressbar2 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
   End
   Begin Project1.simpleprogressbar simpleprogressbar1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
   End
   Begin Project1.simpleprogressbar simpleprogressbar1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
   End
   Begin Project1.simpleprogressbar simpleprogressbar1 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
   End
   Begin Project1.simpleprogressbar simpleprogressbar1 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
   End
   Begin Project1.simpleprogressbar simpleprogressbar1 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
simpleprogressbar2.max = 500
End Sub

Private Sub Timer1_Timer()
simpleprogressbar1(0).value = simpleprogressbar1(0).value + 1
Label1(0).Caption = simpleprogressbar1(0).percent
End Sub

Private Sub Timer2_Timer()
simpleprogressbar1(1).value = simpleprogressbar1(1).value + 1
Label1(1).Caption = simpleprogressbar1(1).percent
End Sub

Private Sub Timer3_Timer()
simpleprogressbar1(2).value = simpleprogressbar1(2).value + 1
Label1(2).Caption = simpleprogressbar1(2).percent
End Sub

Private Sub Timer4_Timer()
simpleprogressbar1(3).value = simpleprogressbar1(3).value + 1
Label1(3).Caption = simpleprogressbar1(3).percent
End Sub

Private Sub Timer5_Timer()
simpleprogressbar1(4).value = simpleprogressbar1(4).value + 1
Label1(4).Caption = simpleprogressbar1(4).percent
End Sub

Private Sub Timer6_Timer()
'total = 0
For d = 0 To 4
total = total + simpleprogressbar1(d).value
DoEvents

Next d

simpleprogressbar2.value = total
Caption = total
Label2.Caption = simpleprogressbar2.percent
End Sub
