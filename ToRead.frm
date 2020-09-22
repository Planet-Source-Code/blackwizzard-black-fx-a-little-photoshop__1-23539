VERSION 5.00
Begin VB.Form ToRead 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "2Read"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "ToRead.frx":0000
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "ToRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
forward Me
End Sub
