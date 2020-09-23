VERSION 5.00
Begin VB.UserControl simpleprogressbar 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   PropertyPages   =   "Copy of UserControl1.ctx":0000
   ScaleHeight     =   255
   ScaleWidth      =   4305
   ToolboxBitmap   =   "Copy of UserControl1.ctx":0014
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   15
         TabIndex        =   1
         Top             =   -120
         Width           =   15
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "simpleprogressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Public max As Long
Attribute max.VB_VarProcData = "PropertyPage1"
Private mvarvalue As Long 'local copy
Private mvarpercent As String 'local copy
Private mvarbackcolor As String 'local copy
Public Property Let backcolor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.percent = 5
    mvarbackcolor = vData
    Picture2.backcolor = vData
End Property


Public Property Get backcolor() As String
Attribute backcolor.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.percent
    backcolor = mvarbackcolor
End Property
Public Property Let percent(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.percent = 5
    mvarpercent = vData
End Property


Public Property Get percent() As String
Attribute percent.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.percent
    percent = mvarpercent
End Property

Private Sub Timer3_Timer()
Label3 = myval
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.ToolTipText = percent
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.ToolTipText = percent
End Sub

Private Sub UserControl_Initialize()
'start with a 0 sized picture to represnt 0%
Picture2.Width = 0
value = 0
' set the progressbar max value like in the common controls version
max = 100
percent = ("00.00")
'value 1
End Sub
Public Function bcolor(color As Long)
On Error Resume Next
Picture2.backcolor = color
End Function


Public Property Let value(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.value = 5
On Error Resume Next
    
 ''set the percentage using a simple math equasion
 
percent = Format(value / (max / 100), "00.00")

'' check if value is higher than the max value so as not to go over 100%
If vData < max Then
''if its lower than 100% set our value
 mvarvalue = vData
Else
'' our value is higher than it should be so setting it to 100%
mvarvalue = max
percent = Format(100, "00.00")
End If

If vData > 0 Then
''set the picture width to the level of the percent for visual purpose
Picture2.Width = (Picture1.Width / 100) * value / (max / 100)
Caption = (Picture1.Width / 100) * value / (max / 100)
If value < max Then
percent = Format(value / (max / 100), "00.00")
Else
percent = Format(100, "00.00")

End If
Else
Exit Property
End If
err:
End Property


Public Property Get value() As Long
Attribute value.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.value
    value = mvarvalue
End Property

Private Sub UserControl_Paint()
Picture2.Width = 0
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
'keep the control the size you want it making it resizable
Picture1.Height = UserControl.Height
Picture1.Width = UserControl.Width
Picture2.Height = UserControl.Height + 80

End Sub
