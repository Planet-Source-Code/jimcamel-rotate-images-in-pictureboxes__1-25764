Attribute VB_Name = "Module1"
'Declare Pi, this is needed for the rotation function
Public Const PI As Double = 3.14159265358979
'GetPixel and SetPixel API. I don't actually use SetPixel
'in this example, Because it doesn't work with pictureboxes
'with autoredraw on (or so it seems)
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
' The function itself. It takes a the parameters of a source picture, the destination picture, the angle, and optional x & y co-ords for the destination image
Public Function RotateSurface(ByRef SourcePicture As PictureBox, ByRef DestPicture As PictureBox, lngAngle As Long, Optional XDest As Long, Optional Ydest As Long)
'declare a bunch of variables
Dim iX As Long, iY As Long
Dim iXDest As Long, iYDest As Long
Dim sngA As Single, SinA As Single, CosA As Single
Dim dblRMax As Long
Dim lngXO As Long, lngYO As Long
Dim lngColor As Long
Dim lWidth As Long, lHeight As Long

'work out the angle in radians
sngA = (360 - lngAngle) * PI / 180
'work out the sine and cosine of the angle (in radians)
SinA = Sin(sngA)
CosA = Cos(sngA)

    
'store the source Image width and image height
lWidth = SourcePicture.Width
lHeight = SourcePicture.Height
'figure out the hypotenuse (diagonal) length of the image
'by using pythagorus's therum.
dblRMax = Sqr(lWidth ^ 2 + lHeight ^ 2)

XDest = XDest + lWidth / 2
Ydest = Ydest + lHeight / 2
'This is the hard maths part. It essentially goes round the source image
'in concentric circles, looks at the colour at each point, and then
'puts it on the destination picture. I did not write much of the code
'below, so I can't explain it that well.
For iX = -dblRMax To dblRMax
    For iY = -dblRMax To dblRMax
    'It takes a while to draw it, so give the system some time.
    DoEvents
        'Figure out where the x and y co-ords are.
        lngXO = lWidth / 2 - (iX * CosA + iY * SinA)
        lngYO = lHeight / 2 - (iX * SinA - iY * CosA)
        'check that the x & y co-ords are 0 or more, and less than the
        'image height.
        If lngXO >= 0 Then
            If lngYO >= 0 Then
                If lngXO < lWidth Then
                    If lngYO < lHeight Then
                        'Use the GetPixel API to get the colour from the
                        'Source Images current X and Y
                        lngColor = GetPixel(SourcePicture.hdc, lngXO, lngYO)
                        'If the colour ain't black
                        If lngColor <> 0 Then
                        'Draw the colour on the destination image.
                        DestPicture.PSet ((lWidth + XDest) - iX, Ydest + iY), lngColor
                        'to use setpixel, unrem the line below and rem the line above
                        'SetPixel is faster than PSet, but has some drawbacks.
                        'SetPixel DestPicture.hdc, XDest + iX, Ydest + iY, lngColor
                        End If
                    End If
                End If
            End If
        End If
    Next iY
Next iX
End Function
