VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form val2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Blur"
   ClientHeight    =   1110
   ClientLeft      =   1530
   ClientTop       =   7290
   ClientWidth     =   7320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Appliquer"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox pic 
      Height          =   1095
      Left            =   4680
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin MSComctlLib.Slider s1 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1085
      _Version        =   393216
      Min             =   1
      Max             =   20
      SelStart        =   2
      TickStyle       =   2
      Value           =   2
   End
End
Attribute VB_Name = "val2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Blur f.p, s1.Value
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
forward Me
pic.Picture = f.p.Image
'Blur pic, 2
End Sub

Private Sub s1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pic.Picture = f.p.Image
Blur pic, s1.Value
End Sub
