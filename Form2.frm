VERSION 5.00
Begin VB.Form f 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Projet en cours"
   ClientHeight    =   4545
   ClientLeft      =   2490
   ClientTop       =   1830
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   15  'Size All
   ScaleHeight     =   4545
   ScaleWidth      =   4095
   Begin VB.OptionButton op2 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton op1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton o 
      Height          =   255
      Index           =   8
      Left            =   8160
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton o 
      Height          =   255
      Index           =   7
      Left            =   8160
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton o 
      Height          =   255
      Index           =   6
      Left            =   8160
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton o 
      Height          =   255
      Index           =   5
      Left            =   8160
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton o 
      Height          =   255
      Index           =   4
      Left            =   8160
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton o 
      Height          =   255
      Index           =   3
      Left            =   8160
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton o 
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton o 
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton o 
      Height          =   255
      Index           =   0
      Left            =   8160
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   284
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2640
         Top             =   1320
      End
      Begin VB.Shape crl 
         Height          =   135
         Left            =   1920
         Shape           =   3  'Circle
         Top             =   2760
         Visible         =   0   'False
         Width           =   135
      End
   End
End
Attribute VB_Name = "f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Dim X1 As Long
Dim Y1 As Long
Dim X2 As Long
Dim Y2 As Long
Dim X3 As Long
Dim Y3 As Long
Dim X4 As Long
Dim Y4 As Long
Dim b As Boolean
Dim i As Long
Dim radius As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then '(+)
radius = radius + 1
ElseIf KeyAscii = 45 Then '(-)
radius = radius - 1
End If
crl.Width = radius * 2
crl.Height = radius * 2
crl.Refresh
menu_FX.SetFocus
End Sub

Private Sub Form_Load()
radius = 5
crl.Width = radius
crl.Height = radius
crl.Refresh
forward Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub p_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then '(+)
radius = radius + 1
ElseIf KeyAscii = 45 Then '(-)
radius = radius - 1
End If
crl.Width = radius * 2
crl.Height = radius * 2
crl.Refresh
menu_FX.SetFocus
End Sub


Private Sub p_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
b = True
X3 = X
Y3 = Y
If o(5).Value = True And b = True Then
find_replace f.p, f.p.Point(X, Y), f.p.ForeColor
End If
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1 = X
Y1 = Y
If o(0).Value = True And b = True Then
p.Line (X2, Y2)-(X, Y)
End If
If o(1).Value = True And b = True Then
p.ForeColor = p.Point(X, Y)
End If
If o(2).Value = True And b = True Then
'''
End If
If o(4).Value = True Then
crl.Left = X - crl.Width / 2
crl.Top = Y - crl.Height / 2
End If
If o(8).Value = True And b = True Then
tagPicture f.p, X, Y, menu_paint.Text1.Text
End If
If op1.Value = True And b = True Then
findRep.Picture1.BackColor = f.p.Point(X, Y)
End If
If op2.Value = True And b = True Then
findRep.Picture2.BackColor = f.p.Point(X, Y)
End If
End Sub

Private Sub p_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'// Basic Paint
If o(3).Value = True And b = True Then
p.Line (X3, Y3)-(X, Y)
End If
'// rectangle
If o(6).Value = True And b = True Then
p.Line (X3, Y3)-(X3, Y)
p.Line (X3, Y3)-(X, Y3)
p.Line (X, Y3)-(X, Y)
p.Line (X3, Y)-(X, Y)
End If
'// Filled Rectangle
If o(7).Value = True Then
p.Line (X3, Y3)-(X3, Y)
p.Line (X3, Y3)-(X, Y3)
p.Line (X, Y3)-(X, Y)
p.Line (X3, Y)-(X, Y)
If X3 < X Then
For i = X3 To X
p.Line (i, Y3)-(i, Y)
Next i
ElseIf X < X3 Then
For i = X To X3
p.Line (i, Y3)-(i, Y)
Next i
End If
End If
'// Circle
If o(4).Value = True And b = True Then
p.Circle (X, Y), radius
End If

If o(5).Value = True And b = True Then
b = False
End If

b = False
X4 = X
Y4 = Y
End Sub

Private Sub Timer1_Timer()
X2 = X1
Y2 = Y1
End Sub
