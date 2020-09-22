VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form menu_paint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tools"
   ClientHeight    =   3465
   ClientLeft      =   6690
   ClientTop       =   2115
   ClientWidth     =   945
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   15
      Top             =   3240
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdcol 
      Left            =   720
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "10"
      Top             =   2925
      Width           =   735
   End
   Begin VB.CommandButton b 
      Height          =   855
      Index           =   8
      Left            =   0
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton b 
      Enabled         =   0   'False
      Height          =   495
      Index           =   13
      Left            =   480
      Picture         =   "Form4.frx":038A
      TabIndex        =   14
      ToolTipText     =   "Rectangle Vide"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton b 
      Height          =   495
      Index           =   12
      Left            =   0
      Picture         =   "Form4.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Rectangle Plein"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton b 
      Enabled         =   0   'False
      Height          =   495
      Index           =   11
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton b 
      Enabled         =   0   'False
      Height          =   495
      Index           =   10
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton b 
      Enabled         =   0   'False
      Height          =   495
      Index           =   9
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton b 
      Height          =   495
      Index           =   7
      Left            =   480
      Picture         =   "Form4.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Rectangle Plein"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton b 
      Height          =   495
      Index           =   6
      Left            =   0
      Picture         =   "Form4.frx":0E28
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Rectangle Vide"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton b 
      Height          =   495
      Index           =   5
      Left            =   480
      Picture         =   "Form4.frx":11B2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cercle Plein"
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton b 
      Height          =   495
      Index           =   4
      Left            =   0
      Picture         =   "Form4.frx":153C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cercle Vide"
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton b 
      Height          =   495
      Index           =   3
      Left            =   480
      Picture         =   "Form4.frx":18C6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ligne"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton b 
      Height          =   495
      Index           =   2
      Left            =   0
      Picture         =   "Form4.frx":1C50
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Loupe"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton b 
      Height          =   495
      Index           =   1
      Left            =   480
      Picture         =   "Form4.frx":1FDA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Pipette"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton b 
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "Form4.frx":2364
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Dessin à la souris"
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "menu_paint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Sub b_Click(Index As Integer)
Select Case Index
Case 0
f.o(0).Value = True
f.o(1).Value = False
f.o(2).Value = False
f.o(3).Value = False
f.o(4).Value = False
f.o(5).Value = False
f.o(6).Value = False
f.o(7).Value = False
f.o(8).Value = False
f.crl.Visible = False
f.p.MousePointer = 2
Case 1
f.o(0).Value = False
f.o(1).Value = True
f.o(2).Value = False
f.o(3).Value = False
f.o(4).Value = False
f.o(5).Value = False
f.o(6).Value = False
f.o(7).Value = False
f.o(8).Value = False
f.crl.Visible = False
Case 2
f.o(0).Value = False
f.o(1).Value = False
f.o(2).Value = True
f.o(3).Value = False
f.o(4).Value = False
f.o(5).Value = False
f.o(6).Value = False
f.o(7).Value = False
f.o(8).Value = False
f.crl.Visible = False
l.p.Picture = f.p.Image
l.Show
Case 3
f.o(0).Value = False
f.o(1).Value = False
f.o(2).Value = False
f.o(3).Value = True
f.o(4).Value = False
f.o(5).Value = False
f.o(6).Value = False
f.o(7).Value = False
f.o(8).Value = False
f.crl.Visible = False
Case 4
f.o(0).Value = False
f.o(1).Value = False
f.o(2).Value = False
f.o(3).Value = False
f.o(4).Value = True
f.o(5).Value = False
f.o(6).Value = False
f.o(7).Value = False
f.o(8).Value = False
f.crl.Visible = True
Case 5
f.o(0).Value = False
f.o(1).Value = False
f.o(2).Value = False
f.o(3).Value = False
f.o(4).Value = False
f.o(5).Value = True
f.o(6).Value = False
f.o(7).Value = False
f.o(8).Value = False
f.crl.Visible = False
Case 6
f.o(0).Value = False
f.o(1).Value = False
f.o(2).Value = False
f.o(3).Value = False
f.o(4).Value = False
f.o(5).Value = False
f.o(6).Value = True
f.o(7).Value = False
f.o(8).Value = False
f.crl.Visible = False
Case 7
f.o(0).Value = False
f.o(1).Value = False
f.o(2).Value = False
f.o(3).Value = False
f.o(4).Value = False
f.o(5).Value = False
f.o(6).Value = False
f.o(7).Value = True
f.o(8).Value = False
f.crl.Visible = False
Case 8
f.o(0).Value = False
f.o(1).Value = False
f.o(2).Value = False
f.o(3).Value = False
f.o(4).Value = False
f.o(5).Value = False
f.o(6).Value = False
f.o(7).Value = False
f.crl.Visible = False
f.o(8).Value = True
Case 12
cdcol.DialogTitle = "choose a color"
cdcol.ShowColor
f.p.ForeColor = cdcol.Color
Picture1.BackColor = cdcol.Color
End Select
End Sub

Private Sub b_KeyPress(Index As Integer, KeyAscii As Integer)
menu_FX.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
menu_FX.SetFocus
End Sub

Private Sub Form_Load()
forward Me
End Sub
