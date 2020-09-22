VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form findRep 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find and replace"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Pic"
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pic"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Annuler"
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Choisir"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choisir"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2400
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   945
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Par celle ci:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remplacer cette couleur:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1785
   End
End
Attribute VB_Name = "findRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cd1.DialogTitle = "choose a color to replace"
cd1.ShowColor
Picture1.BackColor = cd1.color
f.op1.Value = True
End Sub

Private Sub Command2_Click()
cd1.DialogTitle = "choose a color to replace"
cd1.ShowColor
Picture2.BackColor = cd1.color
End Sub

Private Sub Command3_Click()
find_replace f.p, Picture1.BackColor, Picture2.BackColor
f.op1.Value = False
f.op2.Value = False
Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
f.op1.Value = False
f.op2.Value = False
End Sub

Private Sub Command5_Click()
MsgBox "cliquez sur l'image pour selectionner la couleur"
f.op1.Value = True
End Sub

Private Sub Command6_Click()
MsgBox "cliquez sur l'image pour selectionner la couleur"
f.op2.Value = True
End Sub

Private Sub Form_Load()
forward Me
End Sub
