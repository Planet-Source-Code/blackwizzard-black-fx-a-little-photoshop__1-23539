VERSION 5.00
Begin VB.Form val3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Valeur"
   ClientHeight    =   285
   ClientLeft      =   3960
   ClientTop       =   4275
   ClientWidth     =   1440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   1440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "70"
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "val3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Flag f.p, Text1.Text
Unload Me
End Sub

Private Sub Form_Load()
forward Me
End Sub
