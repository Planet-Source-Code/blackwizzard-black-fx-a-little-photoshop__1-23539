VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Black-FX"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "about.frx":0000
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim i As Long
i = 1
Label1.Caption = Label1.Caption & " " & App.Revision
forward Me
End Sub
