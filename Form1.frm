VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form menu_FX 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BlackFX"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":1CCA
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   10
      Left            =   2400
      Picture         =   "Form1.frx":1F14
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   9
      Left            =   2160
      Picture         =   "Form1.frx":215E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   8
      Left            =   1920
      Picture         =   "Form1.frx":23A8
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   7
      Left            =   1680
      Picture         =   "Form1.frx":25F2
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   6
      Left            =   1440
      Picture         =   "Form1.frx":283C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   1200
      Picture         =   "Form1.frx":2A86
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   4
      Left            =   960
      Picture         =   "Form1.frx":2CD0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   3
      Left            =   720
      Picture         =   "Form1.frx":3212
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   2
      Left            =   480
      Picture         =   "Form1.frx":345C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":36A6
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":38F0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   1560
      Width           =   255
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2160
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu file 
      Caption         =   "Fichier"
      Index           =   0
      Begin VB.Menu menu 
         Caption         =   "&New"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu menu 
         Caption         =   "&Open"
         Index           =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu menu 
         Caption         =   "&Save"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu menu 
         Caption         =   "&Quit"
         Index           =   4
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu undoE 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu redo 
         Caption         =   "Redo"
         Shortcut        =   ^W
      End
      Begin VB.Menu restore 
         Caption         =   "Restore"
         Shortcut        =   %{BKSP}
      End
   End
   Begin VB.Menu effect 
      Caption         =   "Effect"
      NegotiatePosition=   3  'Right
      Begin VB.Menu menueffect 
         Caption         =   "Invert"
         Index           =   0
         Shortcut        =   ^I
      End
      Begin VB.Menu menueffect 
         Caption         =   "X-Noise"
         Index           =   1
         Shortcut        =   ^U
      End
      Begin VB.Menu menueffect 
         Caption         =   "Mosaic"
         Index           =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu menueffect 
         Caption         =   "Gaussian Bur"
         Index           =   3
         Shortcut        =   ^B
      End
      Begin VB.Menu menueffect 
         Caption         =   "GrayScale"
         Index           =   4
         Shortcut        =   ^G
      End
      Begin VB.Menu menueffect 
         Caption         =   "luminosité (+)"
         Index           =   5
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu menueffect 
         Caption         =   "luminosité (-)"
         Index           =   6
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu menueffect 
         Caption         =   "STraT"
         Index           =   7
         Shortcut        =   ^R
      End
      Begin VB.Menu menueffect 
         Caption         =   "T°c (Thermique V2)"
         Index           =   8
         Shortcut        =   ^T
      End
      Begin VB.Menu menueffect 
         Caption         =   "Aqua-R"
         Index           =   9
         Shortcut        =   ^A
      End
      Begin VB.Menu menueffect 
         Caption         =   "Photo19"
         Index           =   10
         Shortcut        =   ^P
      End
      Begin VB.Menu menueffect 
         Caption         =   "RGB-FX"
         Index           =   11
         Shortcut        =   ^V
      End
      Begin VB.Menu menueffect 
         Caption         =   "X-Black"
         Index           =   12
         Shortcut        =   ^X
      End
      Begin VB.Menu menueffect 
         Caption         =   "H2O"
         Index           =   13
         Shortcut        =   ^H
      End
      Begin VB.Menu menueffect 
         Caption         =   "PhotoCop"
         Index           =   14
         Shortcut        =   ^L
      End
      Begin VB.Menu menueffect 
         Caption         =   "Comic (encre)"
         Index           =   15
         Shortcut        =   ^E
      End
      Begin VB.Menu menueffect 
         Caption         =   "replace"
         Index           =   16
         Shortcut        =   ^J
      End
      Begin VB.Menu menueffect 
         Caption         =   "Flash"
         Index           =   17
         Shortcut        =   ^C
      End
      Begin VB.Menu menueffect 
         Caption         =   "Flag"
         Index           =   18
         Shortcut        =   ^D
      End
      Begin VB.Menu menueffect 
         Caption         =   "True Noise"
         Index           =   19
         Shortcut        =   ^K
      End
      Begin VB.Menu menueffect 
         Caption         =   "Flip horizontal"
         Index           =   20
         Shortcut        =   ^F
      End
      Begin VB.Menu menueffect 
         Caption         =   "Set WallPaper"
         Index           =   21
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu menueffect 
         Caption         =   "Simple Blur"
         Index           =   22
         Shortcut        =   +{F1}
      End
      Begin VB.Menu menufilter 
         Caption         =   "Filters"
         Begin VB.Menu Fcol 
            Caption         =   "Red"
            Index           =   0
         End
         Begin VB.Menu Fcol 
            Caption         =   "Green"
            Index           =   1
         End
         Begin VB.Menu Fcol 
            Caption         =   "Blue"
            Index           =   2
         End
         Begin VB.Menu Fcol 
            Caption         =   "Yellow"
            Index           =   3
         End
         Begin VB.Menu Fcol 
            Caption         =   "Cyan"
            Index           =   4
         End
         Begin VB.Menu Fcol 
            Caption         =   "Magenta"
            Index           =   5
         End
         Begin VB.Menu Fcol 
            Caption         =   "Any Color"
            Index           =   6
         End
      End
   End
   Begin VB.Menu aboutX 
      Caption         =   "About (?)"
      Begin VB.Menu aboutblack 
         Caption         =   "About Black-FX"
      End
      Begin VB.Menu toread2 
         Caption         =   "To Read"
      End
      Begin VB.Menu author 
         Caption         =   "Author"
      End
      Begin VB.Menu mailme 
         Caption         =   "mail me!"
      End
      Begin VB.Menu sepa01 
         Caption         =   "-"
      End
      Begin VB.Menu bfx 
         Caption         =   "Black-FX"
         Begin VB.Menu bfxontheweb 
            Caption         =   "Black-FX on the web!"
            Begin VB.Menu WebVB 
               Caption         =   "VBFrance (fr)"
            End
            Begin VB.Menu WebPlanet 
               Caption         =   "Planet-source-code (us)"
            End
         End
         Begin VB.Menu downloadBFX 
            Caption         =   "Download the latest version!"
            Begin VB.Menu ZipVB 
               Caption         =   "From VBFrance (fr)"
            End
            Begin VB.Menu ZipPlanet 
               Caption         =   "From Planet-Source-Code (us)"
            End
         End
      End
   End
End
Attribute VB_Name = "menu_FX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

' API stuff for putting bitmaps in menus.
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Const MF_BITMAP = &H4&
Private Const MFT_BITMAP = MF_BITMAP
Private Const MIIM_TYPE = &H10

Dim u As Long

Public Function SetMenuIcon(FrmHwnd As Long, MainMenuNumber As Long, MenuItemNumber As Long, Flags As Long, BitmapUncheckedHandle As Long, BitmapCheckedHandle As Long)
    On Error Resume Next
    Dim lngMenu As Long
    Dim lngSubMenu As Long
    Dim lngMenuItemID As Long
    lngMenu = GetMenu(FrmHwnd)
    lngSubMenu = GetSubMenu(lngMenu, MainMenuNumber)
    lngMenuItemID = GetMenuItemID(lngSubMenu, MenuItemNumber)
    SetMenuIcon = SetMenuItemBitmaps(lngMenu, lngMenuItemID, Flags, BitmapUncheckedHandle, BitmapCheckedHandle)
End Function

Public Sub SetMenuBitmap(ByVal frm As Form, ByVal item_numbers As Variant, ByVal pic As Picture)
Dim menu_handle As Long
Dim i As Integer
Dim menu_info As MENUITEMINFO

    ' Get the menu handle.
    menu_handle = GetMenu(frm.hwnd)
    For i = LBound(item_numbers) To UBound(item_numbers) - 1
        menu_handle = GetSubMenu(menu_handle, item_numbers(i))
    Next i

    ' Initialize the menu information.
    With menu_info
        .cbSize = Len(menu_info)
        .fMask = MIIM_TYPE
        .fType = MFT_BITMAP
        .dwTypeData = pic
    End With

    ' Assign the picture.
    SetMenuItemInfo menu_handle, _
        item_numbers(UBound(item_numbers)), _
        True, menu_info
End Sub

Private Sub aboutblack_Click()
About.Show
End Sub

Private Sub b_Click(Index As Integer)
undo.u(u).Picture = f.p.Image
u = u + 1
If u > 5 Then
u = 5
undo.u(0).Picture = undo.u(1).Image
undo.u(1).Picture = undo.u(2).Image
undo.u(2).Picture = undo.u(3).Image
undo.u(3).Picture = undo.u(4).Image
undo.u(4).Picture = undo.u(5).Image
End If
Select Case Index
Case 0 'new
f.p.Picture = Nothing
Case 1 'save
savepic
Case 2 'open
openimage
Case 3 'invert
Invert f.p
Case 4 'XNoise
Call Blur2
Case 5 'Noise
val1.Show
Case 6 'blur
val2.Show
Case 7 'grayscale
GrayScale f.p
Case 8 'eclaircir
Lighten 15, f.p
Case 9 'assombrir
Darken 15, f.p
Case 10 'STraT
StraT 5, f.p
Case 11 'Vue Thermique V2
Thermique f.p
Case 12 'Aqua-R
Aquarelle f.p, 5
Case 13 'Photo19
Photo f.p
Case 14 'RGB-FX
FX f.p
Case 15 'X-Black
XBlack f.p
End Select
f.p.Refresh
End Sub

Private Sub Command1_Click()
u = u - 1
If u > 5 Then
u = 5
undo.u(0).Picture = undo.u(1).Image
undo.u(1).Picture = undo.u(2).Image
undo.u(2).Picture = undo.u(3).Image
undo.u(3).Picture = undo.u(4).Image
undo.u(4).Picture = undo.u(5).Image
End If
If u < 0 Then u = 0
f.p.Picture = undo.u(u).Image
End Sub

Private Sub Command2_Click()
f.p.Picture = f.p.Picture
End Sub

Private Sub Command3_Click()
About.Show
End Sub

Private Sub Command4_Click()
ToRead.Show
End Sub

Private Sub author_Click()
WeB "http://www.vizue.com/?id=4312", Me.hwnd
End Sub

Private Sub Fcol_Click(Index As Integer)
Select Case Index
Case 0
redfilter f.p
Case 1
greenfilter f.p
Case 2
bluefilter f.p
Case 3
yellowfilter f.p
Case 4
cyanfilter f.p
Case 5
Magentafilter f.p
Case 6
cd1.ShowColor
Colorfilter f.p, cd1.color
End Select
f.p.Refresh
End Sub

Private Sub Form_Load()
Me.Height = 600
Me.Width = 2760
Me.Left = 3105
Me.Top = 1215
menu_paint.Top = 1215
menu_paint.Left = 6645
f.Top = 1830
f.Left = 2490
Dim i As Long
For i = 0 To menueffect.Count
'SetMenuBitmap Me, Array(2, i), Me.Picture
SetMenuIcon Me.hwnd, 2, i, 0, Me.Picture, Me.Picture
Next i
SetMenuIcon Me.hwnd, 0, 0, 0, Picture1(0).Picture, Me.Picture
SetMenuIcon Me.hwnd, 0, 1, 0, Picture1(3).Picture, Me.Picture
SetMenuIcon Me.hwnd, 3, 0, 0, Picture1(1).Picture, Me.Picture
SetMenuIcon Me.hwnd, 3, 1, 0, Picture1(2).Picture, Me.Picture
SetMenuIcon Me.hwnd, 0, 3, 0, Picture1(5).Picture, Me.Picture
SetMenuIcon Me.hwnd, 0, 2, 0, Picture1(6).Picture, Me.Picture
SetMenuIcon Me.hwnd, 3, 2, 0, Picture1(7).Picture, Me.Picture
SetMenuIcon Me.hwnd, 1, 0, 0, Picture1(8).Picture, Me.Picture
SetMenuIcon Me.hwnd, 1, 1, 0, Picture1(9).Picture, Me.Picture
SetMenuIcon Me.hwnd, 1, 2, 0, Picture1(10).Picture, Me.Picture
SetMenuIcon Me.hwnd, 3, 3, 0, Me.Picture, Me.Picture
f.Show
menu_paint.Show
'menu_FX.Left = 50
'f.Left = menu_FX.Left + menu_FX.Width + 20
'menu_paint.Left = f.Left + f.Width + 20
u = 0
forward Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Public Sub openimage()
cd1.DialogTitle = "ouvrir une image"
cd1.Filter = "All Suported|*.jpg;*.jpeg;*.gif;*.bmp;*.rle|gif|*.gif|Bitmap|*.bmp;*.rle|jpg|*.jpg;*.jpeg"
cd1.ShowOpen
If cd1.FileName <> "" Then
f.p.Picture = LoadPicture(cd1.FileName)
f.Width = f.p.Width
f.Height = f.p.Height + 255
End If
End Sub

Public Sub savepic()
cd1.DialogTitle = "Save picture"
cd1.Filter = "Bitmap (*.bmp)|*.bmp|rle Bitmap (*.rle)|*.rle"
cd1.ShowSave
If cd1.FileName <> "" Then
Form2.Picture1.Picture = f.p.Image
SavePicture Form2.Picture1, cd1.FileName
End If
End Sub

Private Sub mailme_Click()
WeB "mailto:blackwizzard@wanadoo.fr", Me.hwnd
End Sub

Private Sub menu_Click(Index As Integer)
Select Case Index
Case 1
f.p.Picture = Nothing
f.p.Width = 4095
f.Width = 4095
f.p.Height = 4290
f.Height = 4545
Case 2
openimage
Case 3
savepic
Case 4
End
End Select
End Sub

Private Sub menueffect_Click(Index As Integer)
On Error Resume Next
undo.u(u).Picture = f.p.Image
u = u + 1
If u > 5 Then
u = 5
undo.u(0).Picture = undo.u(1).Image
undo.u(1).Picture = undo.u(2).Image
undo.u(2).Picture = undo.u(3).Image
undo.u(3).Picture = undo.u(4).Image
undo.u(4).Picture = undo.u(5).Image
End If
Select Case Index
Case 0 'invert
Invert f.p
Case 1 'XNoise
Call Blur2
Case 2 'Mosaic
val1.Show
Case 3 'blur
val2.Show
Case 4 'gray
GrayScale f.p
Case 5 'L+
Lighten 15, f.p
Case 6 'L-
Darken 15, f.p
Case 7 'STraT
StraT 5, f.p
Case 8 'T°c
Thermique f.p
Case 9 'Aqua-R
Aquarelle f.p, 5
Case 10 'Photo19
Photo f.p
Case 11 'RGBFX
FX f.p
Case 12 'Xblack
XBlack f.p
Case 13 'H2O
H2O f.p
Case 14 'photocop
photocop f.p
Case 15 'comic
BD f.p, 1, 10
Case 16
findRep.Show
Case 17
flash f.p, 1, 10
Case 18
val3.Show
Case 19
Noise f.p, 500
Case 20
Flip_Horizontal f.p
Case 21
Form2.Picture1.Picture = f.p.Image
SavePicture Form2.Picture1, "c:\windows\FXpaper.bmp"
Call SystemParametersInfo(20, 1, "c:\windows\FXpaper.bmp", 1)
Case 22
Blur2
End Select
f.p.Refresh
undo.u(u).Picture = f.p.Image
undoE.Enabled = True
End Sub

Private Sub redo_Click()
u = u + 1
If u > 5 Then
u = 5
undo.u(1).Picture = undo.u(0).Image
undo.u(2).Picture = undo.u(1).Image
undo.u(3).Picture = undo.u(2).Image
undo.u(4).Picture = undo.u(3).Image
undo.u(5).Picture = undo.u(4).Image
redo.Enabled = False
End If
If u < 0 Then u = 0: undoE.Enabled = False
f.p.Picture = undo.u(u).Image
End Sub

Private Sub restore_Click()
f.p.Picture = f.p.Picture
End Sub


Private Sub toread2_Click()
ToRead.Show
End Sub

Private Sub undoE_Click()
redo.Enabled = True
u = u - 1
If u > 5 Then
u = 5
undo.u(0).Picture = undo.u(1).Image
undo.u(1).Picture = undo.u(2).Image
undo.u(2).Picture = undo.u(3).Image
undo.u(3).Picture = undo.u(4).Image
undo.u(4).Picture = undo.u(5).Image
End If
If u < 0 Then u = 0: undoE.Enabled = False
f.p.Picture = undo.u(u).Image
End Sub

Private Sub WebPlanet_Click()
WeB "http://www.planet-source-code.com/xq/ASP/txtCodeId.23539/lngWId.1/qx/vb/scripts/ShowCode.htm", Me.hwnd
End Sub

Private Sub WebVB_Click()
WeB "http://www.vbfrance.com/article.asp?Val=1368", Me.hwnd
End Sub

Private Sub ZipPlanet_Click()
WeB "http://www.planet-source-code.com/upload/ftp/Black-FX201925292001.zip", Me.hwnd
End Sub

Private Sub ZipVB_Click()
WeB "http://www.vbfrance.com/fichier.asp?Val=1368&F=W", Me.hwnd
End Sub
