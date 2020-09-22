VERSION 5.00
Begin VB.Form l 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Loupe"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image p 
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "l"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
forward Me
End Sub

Private Sub Form_Resize()
p.Width = Me.Width
p.Height = Me.Height - 20
End Sub
