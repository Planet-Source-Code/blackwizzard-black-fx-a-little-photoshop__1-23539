Attribute VB_Name = "image32"
Option Explicit
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Global rRed As Long, rBlue As Long, rGreen As Long
Global rRed2 As Long, rBlue2 As Long, rGreen2 As Long
Global CCr As Long, CCg As Long, CCb As Long
Global Coef As Long

Public Function RGBfromLONG(LongCol As Long)
' Get The Red, Blue And Green Values Of A Colour From The Long Value
Dim Blue As Double, Green As Double, Red As Double, GreenS As Double, BlueS As Double
Blue = Fix((LongCol / 256) / 256)
Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))
rRed = Red: rBlue = Blue: rGreen = Green
End Function

Public Function RGBfromLONG2(LongCol As Long)
' Get The Red, Blue And Green Values Of A Colour From The Long Value
Dim Blue As Double, Green As Double, Red As Double, GreenS As Double, BlueS As Double
Blue = Fix((LongCol / 256) / 256)
Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))
rRed = Red: rBlue = Blue: rGreen = Green
End Function

Public Function GetRandomNumber(Upper As Integer, Lower As Integer) As Integer
'Get a random number
Randomize
GetRandomNumber = Int((Upper) * Rnd)
End Function

Sub Noise(picBox As PictureBox, Intensity As Integer)
'Add noise to a picture
Dim X As Integer, W As Integer, h As Integer, Num As Integer, Num2 As Integer
picBox.ScaleMode = 3
W = picBox.ScaleWidth
h = picBox.ScaleHeight
For X = 1 To Intensity * 50
    Randomize
    Num = Int(Rnd * W - 1) + 1
    Randomize
    Num2 = Int(Rnd * h - 1) + 1
    SetPixel picBox.hdc, Num, Num2, GetPixel(picBox.hdc, Num2, Num)
Next X
End Sub
Sub Mosaic(picBox As PictureBox, Size As Integer)
On Error Resume Next
'Pixelate a picture
Dim W As Integer, h As Integer, NumC As Integer
Dim color As Long, CA As Integer
Dim C(1 To 100) As Long, s As Integer
Dim g As Long, r As Long, b As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight - 2 Step Size
    For W = 0 To picBox.ScaleWidth - 2 Step Size
        NumC = 1
        For s = 1 To Size
            C(NumC) = GetPixel(picBox.hdc, W, h)
            NumC = NumC + 1
            C(NumC) = GetPixel(picBox.hdc, W + s, h)
            NumC = NumC + 1
            C(NumC) = GetPixel(picBox.hdc, W + s, h + s)
            NumC = NumC + 1
            C(NumC) = GetPixel(picBox.hdc, W, h + s)
            NumC = NumC + 1
        Next s
        For CA = 1 To NumC
            RGBfromLONG C(CA)
            g = g + rGreen
            r = r + rRed
            b = b + rBlue
        Next CA
        r = r / NumC
        g = g / NumC
        b = b / NumC
        color = RGB(r, g, b)
        
        For s = 0 To Size
            picBox.Line (W + s, h)-(W + s, h + Size), color, BF
        Next s
    Next W
    DoEvents
Next h
End Sub
Sub Lighten(percent As Integer, picBox As PictureBox)
'Lighten a picture

Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long

newVal = percent * 5
picBox.ScaleMode = 3

For W = 0 To picBox.ScaleWidth
    For h = 0 To picBox.ScaleHeight
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        opRed = rRed
        opGreen = rGreen
        opBlue = rBlue
        rRed = rRed + newVal
        If rRed > -1 And rRed < 256 Then opRed = rRed
        
        rGreen = rGreen + newVal
        If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
        rBlue = rBlue + newVal
        If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
        If rRed <> 1000 Then
           C = RGB(opRed, opGreen, opBlue)
           SetPixel picBox.hdc, W, h, C
        End If
    Next h
    picBox.Refresh
Next W
End Sub


Sub Darken(percent As Integer, picBox As PictureBox)
'Darken a picture
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim icRed As Long, icBlue As Long, icGreen As Long
Dim opRed As Long, opBlue As Long, opGreen As Long

newVal = percent * -5
picBox.ScaleMode = 3

For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        opRed = rRed
        opBlue = rBlue
        opGreen = rGreen
        rRed = rRed + newVal
        If rRed > -1 And icRed < 256 Then opRed = rRed
        
        rGreen = rGreen + newVal
        If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
        rBlue = rBlue + newVal
        If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
        If rRed <> 1000 Then
            If opRed < 0 Then opRed = 0
            If opGreen < 0 Then opGreen = 0
            If opBlue < 0 Then opBlue = 0
           C = RGB(opRed, opGreen, opBlue)
           SetPixel picBox.hdc, W, h, C
        End If
    Next W
    picBox.Refresh
Next h
End Sub


Public Sub GrayScale(picBox As PictureBox)
'Turn a color image to greyscale
Dim AveCol As Integer, A As Integer
Dim Y As Integer, X As Integer

picBox.ScaleMode = 3
For Y = 0 To picBox.ScaleHeight
    For X = 0 To picBox.ScaleWidth
        AveCol = 0
        A = 0
        RGBfromLONG GetPixel(picBox.hdc, X, Y)
        AveCol = AveCol + rGreen: A = A + 1
        If AveCol <= 0 Then AveCol = 0
        AveCol = (AveCol / A)
        SetPixel picBox.hdc, X, Y, RGB(AveCol, AveCol, AveCol)
    Next X
    picBox.Refresh
Next Y
End Sub


Function LightenPixel(pixelLong As Long, percent As Integer)
'Lighten only one pixel
Dim newVal As Integer, C As Long, opRed As Long, opGreen As Long, opBlue As Long
newVal = percent * 5
C = pixelLong
RGBfromLONG C
opRed = rRed
opGreen = rGreen
opBlue = rBlue
rRed = rRed + newVal
If rRed > -1 And rRed < 256 Then opRed = rRed

rGreen = rGreen + newVal
If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
rBlue = rBlue + newVal
If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
If rRed <> 1000 Then
    C = RGB(opRed, opGreen, opBlue)
    LightenPixel = C
End If
End Function


Function DarkenPixel(pixelLong As Long, percent As Integer) As Long
'Darken only one pixel
Dim newVal As Integer, C As Long, opRed As Long, opGreen As Long, opBlue As Long
newVal = percent * -5
C = pixelLong
RGBfromLONG C
opRed = rRed
opGreen = rGreen
opBlue = rBlue
rRed = rRed + newVal
If rRed > -1 And rRed < 256 Then opRed = rRed

rGreen = rGreen + newVal
If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
rBlue = rBlue + newVal
If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
If rRed <> 1000 Then
    C = RGB(opRed, opGreen, opBlue)
    DarkenPixel = C
End If
End Function


Sub Blur(picBox As PictureBox, Intensity As Integer)
'Blur a picture
On Error Resume Next
Dim W As Integer, h As Integer, NumC As Integer
Dim color As Long, CA As Integer, Size As Integer, Size2 As Integer
Dim C(1 To 100) As Long, s As Integer, i As Integer
Dim g As Long, r As Long, b As Long
picBox.ScaleMode = 3
Size = 0.5
Size2 = 0.5

For i = 1 To Intensity
    For W = 0 To picBox.ScaleWidth - 2 Step Size
        For h = 0 To picBox.ScaleHeight - 2 Step Size2
            NumC = 1
            For s = 1 To Size
                C(NumC) = GetPixel(picBox.hdc, W, h)
                NumC = NumC + 1
                C(NumC) = GetPixel(picBox.hdc, W + s, h)
                NumC = NumC + 1
                C(NumC) = GetPixel(picBox.hdc, W + s, h + s)
                NumC = NumC + 1
                C(NumC) = GetPixel(picBox.hdc, W, h + s)
                NumC = NumC + 1
            Next s
            For CA = 1 To NumC
                RGBfromLONG C(CA)
                g = g + rGreen
                r = r + rRed
                b = b + rBlue
            Next CA
            If g > 0 And r > 0 And b > 0 Then
                r = r / NumC
                g = g / NumC
                b = b / NumC
            Else
                r = 0
                g = 0
                b = 0
            End If
            color = RGB(r, g, b)
            
            For s = 0 To Size
                picBox.Line (W + s, h)-(W + s, h + Size), color, BF
            Next s
        Next h
       ' DoEvents
        picBox.Refresh
    Next W
   ' DoEvents
Next i
End Sub

Function InvertPixel(colorLong As Long) As Long
'Invert the color of a pixel
Dim opRed As Long, opGreen As Long, opBlue As Long
RGBfromLONG colorLong

InvertPixel = RGB(255 - rRed, 255 - rGreen, 255 - rBlue)
End Function
Sub Invert(picBox As PictureBox)
'Invert the image of a picturebox
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long

picBox.ScaleMode = 3

For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth

        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        opRed = 255 - rRed
        opGreen = 255 - rGreen
        opBlue = 255 - rBlue
        C = RGB(opRed, opGreen, opBlue)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Sub Flip_Horizontal(picBox As PictureBox)
Dim W As Integer, h As Integer, Num As Integer
Dim OldColor(0 To 1000, 0 To 1000) As Long, cColor As Long
picBox.ScaleMode = 3
Num = picBox.ScaleWidth / 2
For W = picBox.ScaleWidth / 2 To picBox.ScaleWidth
    For h = 0 To picBox.ScaleHeight
        cColor = GetPixel(picBox.hdc, Num, h)
        OldColor(W - (picBox.ScaleWidth / 2), h) = GetPixel(picBox.hdc, W, h)
        SetPixel picBox.hdc, W, h, cColor
    Next h
    Num = Num - 1
    DoEvents
Next W


Num = picBox.ScaleWidth / 2
For W = 0 To picBox.ScaleWidth / 2
    For h = 0 To picBox.ScaleHeight
        SetPixel picBox.hdc, Num, h, OldColor(W, h)
    Next h
    Num = Num - 1
    DoEvents
Next W
End Sub

Public Sub StraT(percent As Integer, picBox As PictureBox)
Dim s As Long
Dim newVal As Integer, h As Integer, W As Integer, K As Integer, newVal2 As Integer
Dim C As Long
Dim icRed As Long, icBlue As Long, icGreen As Long
Dim opRed As Long, opBlue As Long, opGreen As Long
s = 0
newVal = percent * -5
picBox.ScaleMode = 3

For h = 0 To picBox.ScaleHeight 'Step 2
s = s + 1
If s = 5 Then s = 0
If s = 0 Then newVal2 = newVal / 5
If s = 1 Then newVal2 = newVal / 2
If s = 2 Then newVal2 = newVal
If s = 3 Then newVal2 = newVal / 2
If s = 4 Then newVal2 = newVal / 5
    For W = 0 To picBox.ScaleWidth 'Step 5
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        opRed = rRed
        opBlue = rBlue
        opGreen = rGreen
        rRed = rRed + newVal2
        If rRed > -1 And icRed < 256 Then opRed = rRed
        rGreen = rGreen + newVal2
        If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
        rBlue = rBlue + newVal2
        If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
        If rRed <> 1000 Then
            If opRed < 0 Then opRed = 0
            If opGreen < 0 Then opGreen = 0
            If opBlue < 0 Then opBlue = 0
           C = RGB(opRed, opGreen, opBlue)
           SetPixel picBox.hdc, W, h, C
           End If
    Next W
    picBox.Refresh
Next h
End Sub

Sub Thermique(picBox As PictureBox)
'Invert the image of a picturebox
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        If rRed > rGreen And rRed > rBlue Then
        opRed = rRed
        opGreen = 0
        opBlue = 0
        ElseIf rGreen > rRed And rGreen > rBlue Then
        opRed = 0
        opGreen = rGreen
        opBlue = 0
        ElseIf rBlue > rGreen And rBlue > rRed Then
        opRed = 0
        opGreen = 0
        opBlue = rBlue
        End If
        C = RGB(opRed, opGreen, opBlue)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Function Lp(Xp, Yp, percent As Integer, picBox As PictureBox)
'Lighten only one pixel
Dim newVal As Integer, C As Long, opRed As Long, opGreen As Long, opBlue As Long
newVal = percent * 5
C = picBox.Point(Xp, Yp)
RGBfromLONG C
opRed = rRed
opGreen = rGreen
opBlue = rBlue
rRed = rRed + newVal
If rRed > -1 And rRed < 256 Then opRed = rRed
rGreen = rGreen + newVal
If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
rBlue = rBlue + newVal
If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
If rRed <> 1000 Then
    C = RGB(opRed, opGreen, opBlue)
    Lp = C
    SetPixel picBox.hdc, Xp, Yp, C
End If
End Function

Function Dp(Xp, Yp, percent As Integer, picBox As PictureBox) As Long
'Darken only one pixel
On Error Resume Next
Dim newVal As Integer, C As Long, opRed As Long, opGreen As Long, opBlue As Long
newVal = percent * -5
C = picBox.Point(Xp, Yp)
RGBfromLONG C
opRed = rRed
opGreen = rGreen
opBlue = rBlue
rRed = rRed + newVal
If rRed > -1 And rRed < 256 Then opRed = rRed
rGreen = rGreen + newVal
If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
rBlue = rBlue + newVal
If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
If rRed <> 1000 Then
    C = RGB(opRed, opGreen, opBlue)
    Dp = C
    SetPixel picBox.hdc, Xp, Yp, C
End If
End Function

Public Sub Aquarelle(picBox As PictureBox, valR As Integer)
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        If rRed > 120 And rGreen > 120 And rBlue > 120 Then
        Dp W, h, valR, picBox
        ElseIf rRed < 120 And rGreen < 120 And rBlue < 120 Then
        Lp W, h, valR, picBox
        End If
        C = RGB(opRed, opGreen, opBlue)
'        SetPixel PicBox.hdc, W, h, C
'        Dp W, h, 5, PicBox
'        Lp W, h, 5, PicBox
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub Photo(picBox As PictureBox)
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long
Dim AveCol As Integer, A As Integer
Dim Y As Integer, X As Integer
picBox.ScaleMode = 3
GrayScale picBox
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        If rRed > 120 And rGreen > 120 And rBlue > 120 Then
        Lp W, h, 15, picBox
        ElseIf rRed < 120 And rGreen < 120 And rBlue < 120 And rRed > 50 And rGreen > 50 And rBlue > 50 Then
        Lp W, h, 15, picBox
        ElseIf rBlue > rGreen And rBlue > rRed Then
        opRed = 0
        opGreen = 0
        opBlue = rBlue
        End If
        C = RGB(opRed, opGreen, opBlue)
'        SetPixel PicBox.hdc, W, h, C
'        Dp W, h, 5, PicBox
'        Lp W, h, 5, PicBox
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub FX(picBox As PictureBox)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long, C2 As Long
Dim opRed As Long, opBlue As Long, opGreen As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        opRed = rRed
        opGreen = rGreen
        opBlue = rBlue
        C = RGB(opBlue, opRed, opGreen)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub XBlack(picBox As PictureBox)
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        If rRed > (255 / 2) And rGreen > (255 / 2) And rBlue > (255 / 2) Then
        opRed = 0
        opGreen = 0
        opBlue = 0
        ElseIf rRed < (255 / 2) And rGreen < (255 / 2) And rBlue < (255 / 2) Then
        opRed = 255
        opGreen = 255
        opBlue = 255
        End If
        C = RGB(opRed, opGreen, opBlue)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub Invert2()
Dim i As Long
Dim j As Long, pixvalue, red_val, green_val, blue_val, hex_pixval

For i = 0 To f.p.ScaleHeight - 1
    For j = 0 To f.p.ScaleWidth - 1
      pixvalue = f.p.Point(j, i)
      hex_pixval = Hex(pixvalue)
      red_val = "&h" & Mid$(hex_pixval, 5, 2)
      green_val = "&h" & Mid$(hex_pixval, 3, 2)
      blue_val = "&h" & Mid$(hex_pixval, 1, 2)

      If red_val = "&h" Then red_val = "&h0"
      If green_val = "&h" Then green_val = "&h0"
      If blue_val = "&h" Then blue_val = "&h0"

      red_val = &HFF - red_val
      green_val = &HFF - green_val
      blue_val = &HFF - blue_val

      f.p.PSet (j, i), RGB(red_val, green_val, blue_val)
      'f.p.Refresh
    Next
Next
End Sub

Public Sub Blur2()
Dim i As Long
Dim j As Long, pixval, red_val, green_val, blue_val

For i = 1 To f.p.ScaleHeight - 3
    For j = 1 To f.p.ScaleWidth - 3
        pixval = f.p.Point(j + CInt(Rnd * 10), i + CInt(Rnd * 10))
        red_val = "&h" & Mid$(CStr(Hex(pixval)), 5, 2)
        green_val = "&h" & Mid$(CStr(Hex(pixval)), 3, 2)
        blue_val = "&h" & Mid$(CStr(Hex(pixval)), 1, 2)
        If red_val = "&h" Then red_val = "&h0"
        If green_val = "&h" Then green_val = "&h0"
        If blue_val = "&h" Then blue_val = "&h0"
        f.p.PSet (j, i), RGB(red_val, green_val, blue_val)
    Next
    f.p.Refresh
Next
End Sub

Public Sub BD(picBox As PictureBox, valR As Integer, valR2 As Integer)
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
    For K = 0 To valR2
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        If rRed > 120 And rGreen > 120 And rBlue > 120 Then
        Lp W, h, valR, picBox
        ElseIf rRed < 120 And rGreen < 120 And rBlue < 120 Then
        Dp W, h, valR, picBox
        End If
        C = RGB(opRed, opGreen, opBlue)
        Next K
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub H2O(picBox As PictureBox)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long, z As Long, v As Long
Dim opRed As Long, opBlue As Long, opGreen As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
z = z + 1
'If z = 30 Then z = 0
v = Sin(z) * 5
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        opRed = rRed
        opGreen = rGreen
        opBlue = rBlue
        C = RGB(opRed, opGreen, opBlue)
        SetPixel picBox.hdc, W, h - v, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub photocop(picBox As PictureBox)
Dim AveCol As Integer, A As Integer
Dim Y As Integer, X As Integer

picBox.ScaleMode = 3
For Y = 0 To picBox.ScaleHeight
    For X = 0 To picBox.ScaleWidth
        AveCol = 0
        A = 0
        RGBfromLONG GetPixel(picBox.hdc, X, Y)
        AveCol = AveCol + rGreen: A = A + 1
        If AveCol <= 0 Then AveCol = 0
        If AveCol < (255) / 2 Then AveCol = AveCol / 5
        If AveCol > (255) / 2 Then AveCol = AveCol * 5
        AveCol = (AveCol / A)
        SetPixel picBox.hdc, X, Y, RGB(AveCol, AveCol, AveCol)
    Next X
    picBox.Refresh
Next Y
End Sub


Public Sub tagPicture(picBox As PictureBox, X As Single, Y As Single, Val As Long)
On Error Resume Next
Dim i
For i = 1 To Val
'Dp (X + Sin(i) * i), (Y + Cos(i) * i), 50, PicBox
SetPixel picBox.hdc, (X + Sin(i) * i), (Y + Cos(i) * i), picBox.ForeColor
picBox.Refresh
Next i
End Sub

Public Sub find_replace(picBox As PictureBox, color1 As OLE_COLOR, color2 As OLE_COLOR)
Dim AveCol As Integer, A As Integer
Dim Y As Integer, X As Integer
picBox.ScaleMode = 3
For Y = 0 To picBox.ScaleHeight
    For X = 0 To picBox.ScaleWidth
    If picBox.Point(X, Y) = color1 Then
    SetPixel picBox.hdc, X, Y, color2
    End If
    Next X
    picBox.Refresh
Next Y
End Sub

Public Sub Flag(picBox As PictureBox, Val As Long)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long, z As Long, v As Long
Dim opRed As Long, opBlue As Long, opGreen As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
z = z + 1
'If z = 30 Then z = 0
v = Sin(z / Int(picBox.ScaleHeight / 3.14961)) * Val
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        opRed = rRed
        opGreen = rGreen
        opBlue = rBlue
        C = RGB(opRed, opGreen, opBlue)
        SetPixel picBox.hdc, W - v, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub flash(picBox As PictureBox, valR As Integer, valR2 As Integer)
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long, op As Long
picBox.ScaleMode = 3
'GrayScale picBox
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
    For K = 0 To valR2
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        CoefCol rRed, rGreen, rBlue
        op = Coef
        C = RGB(op, op, op)
        SetPixel picBox.hdc, W, h, C
        Next K
    Next W
    picBox.Refresh
Next h
End Sub

Public Function CoefCol(r As Long, g As Long, b As Long)
Coef = (r + g + b) / 2.45
End Function

Public Sub redfilter(picBox As PictureBox)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long, op As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        C = RGB(rRed, 0, 0)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub greenfilter(picBox As PictureBox)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long, op As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        C = RGB(0, rGreen, 0)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub bluefilter(picBox As PictureBox)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long, op As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        C = RGB(0, 0, rBlue)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub cyanfilter(picBox As PictureBox)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long, op As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        C = RGB(0, rGreen, rBlue)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub Magentafilter(picBox As PictureBox)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long, op As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        C = RGB(rRed, 0, rBlue)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub yellowfilter(picBox As PictureBox)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long, op As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        RGBfromLONG C
        C = RGB(rRed, rGreen, 0)
        SetPixel picBox.hdc, W, h, C
    Next W
    picBox.Refresh
Next h
End Sub

Public Sub Colorfilter(picBox As PictureBox, color As Long)
On Error Resume Next
Dim newVal As Integer, h As Integer, W As Integer, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long, op As Long
picBox.ScaleMode = 3
For h = 0 To picBox.ScaleHeight
    For W = 0 To picBox.ScaleWidth
        C = GetPixel(picBox.hdc, W, h)
        'RGBfromLONG C
        'RGBfromLONG2 color
        'opRed = rRed Mod rRed2
        'opGreen = rGreen Mod rGreen2
        'opBlue = rBlue Mod rBlue2
        'C = RGB(opRed, opGreen, opBlue)
        SetPixel picBox.hdc, W, h, C Mod color / 5
    Next W
    picBox.Refresh
Next h
Exit Sub
End Sub
