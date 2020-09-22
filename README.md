<div align="center">

## Mosaic


</div>

### Description

Takes a picturebox, and it's contents, and runs an animated mosaic transition through it
 
### More Info
 
pctMosaic, the picturebox object that you're wanting to manipulate

MosaicMode, set it to 1 for mosaic, 2 for demosaic, 3 for mosaic, then demosaic

Nothing. If you want to edit it, that's another story :)

Can crash some computers. Seems to be a display driver to windows problem! I don't know what causes this!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dave Hng](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-hng.md)
**Level**          |Unknown
**User Rating**    |6.0 (600 globes from 100 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dave-hng-mosaic__1-883/archive/master.zip)

### API Declarations

```
'Functions for Processing Bitmaps
Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type
Type SafeArrayBound
  cElements As Long
  lLbound As Long
End Type
Type SafeArray2D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  bounds(0 To 1) As SafeArrayBound
End Type
```


### Source Code

```
Sub GenMosaic(pctMosaic As Variant, MosaicMode As Integer)
'Mosaic Mode is 1 for Mosaic, 2 for DeMosaic
'Declare all objects
'======================================================================
'  This code is (C) StarFox / Dave Hng '98
'
'  Posted on http://www.planet-source-code.com during May '98.
'
'  If you distribute this code, make sure that the complete listing is intact, with these
'  comments! If you use it in a program, don't worry about this introduction.
'
'  Email: StarFox@earthcorp.com or psychob@inf.net.au
'  UIN: 866854
'
'  Please credit me if you use this code! As far as i know, this is the only nice(ish) VB
'  image manip sub that i've seen! This is one major code hack! :)
'
'  Takes a picturebox, and runs a animated mosaic transition though it!
'
'  Uses Safearrays, CopyMemory, Bitmap basics. Not for the faint hearted.
'
'  pctMosaic is a picturebox object that you want to run the transition through
'  MosaicMode is an integer, indicating what steps of the mosaic you want to run though.
'  1 is mosaic up, 2 is mosaic down, 3 is mosaic up, then down again. Experiment!
'
'  Not very efficient, but the code runs at about 2x to 10x emulated speed when compiled to
'  native code! It runs really really fast compiled under the native code compiler!
'  It's capable of animating a small bitmap on a 486dx2/80, with the interval set to 1, and
'  no re-redraws.
'
'  Only works on 256 colour, single plane bitmaps. I'll write one for truecolour images when
'  i figure out how the RGBQuad type works, (Can anyone help?) and i've finished high school.
'
'  You can change the for.. next statements with the K and L variables to change the speed of
'  the function. K is the mosaic depth, L is the number of times to call the function (limits
'  speed, so you can see it better)
'
'  Thanks to the guys that wrote the VBPJ article on direct access to memory. Without that info
'  or ideas, i wouldnt've been able to write the sub.
'
'  This code is used in StarLaunch, my multi emulator launcher:
'  http://starlaunch.home.ml.org
'  As a transition for screen size previews for snes emulators.
'
'  Note: It does crash some computers, for no known reason.
'  I think it's as video card -> video driver problem.
'  Don't break while this sub is running, unless you really have to. If you want to stop
'  execution, you must call the cleanup code associated with what the sub's doing.
'  (Copymemory the pointer to 0& again)
'
'  Have fun!
'
'  "If you think it's not possible, make it!"
'
'  -StarFox
Static mosaicgoing As Boolean
'Keep a static variable to check if the sub's running. If it is, EXIT! Otherwise, GPF!
If mosaicgoing = True Then Exit Sub
mosaicgoing = True
'Init variables
Dim pict() As Byte
Dim SA As SafeArray2D, bmp As BITMAP
Dim r As Integer, c As Integer, Value As Byte, i As Integer, colour As Integer, j As Integer, k As Integer, L As Integer
Dim pCenter As Integer, pC As Integer, pR As Integer
Dim rRangei As Integer, rRangej As Integer, ti As Integer, ti2 As Integer
Dim uC As Integer, uR As Integer
Dim PictureArray() As Byte
Dim mRange As Integer
Dim cLimit As Integer, rLimit As Integer
'Copy to the array
'======================================================================
GetObjectAPI pctMosaic.Picture, Len(bmp), bmp
If bmp.bmPlanes <> 1 Or bmp.bmBitsPixel <> 8 Then
  MsgBox "Non-256 colour bitmap detected. No mosaic effects"
  Exit Sub
End If
'Init the SafeArray
With SA
  .cbElements = 1
  .cDims = 2
  .bounds(0).lLbound = 0
  .bounds(0).cElements = bmp.bmHeight
  .bounds(1).lLbound = 0
  .bounds(1).cElements = bmp.bmWidthBytes
  .pvData = bmp.bmBits
End With
'Map the pointer over
CopyMemory ByVal VarPtrArray(pict), VarPtr(SA), 4
'Make a temporary array to hold the bitmap data.
ReDim PictureArray(UBound(pict, 1), UBound(pict, 2))
'Copy the bitmap into this array. I could use copymemory again, but this is fast enough,
'and a lot safer :)
For c = 0 To UBound(pict, 1)
  For r = 0 To UBound(pict, 2)
      PictureArray(c, r) = pict(c, r)
  Next r
Next c
'Clean up
CopyMemory ByVal VarPtrArray(pict), 0&, 4
'======================================================================
Select Case MosaicMode
  Case 1
  'Mosaic
    For k = 1 To 16 Step 1
      For L = 1 To 1
		'Cube roots used, because the squaring effect looks nicer. Also, due to the
		'Nature of my code, it hides irregular the pixel expansion
        mRange = k ^ 1.5
        GoSub Mosaic
      Next L
    Next k
  Case 2
  'DeMosaic
    For k = 16 To 0 Step -(1)
      For L = 1 To 1
        mRange = k ^ 1.5
        GoSub Mosaic
      Next L
    Next k
  Case 3
  'Mosaic, then DeMosaic
    For k = 1 To 8 Step 1
      mRange = k ^ 1.5
        GoSub Mosaic
    Next k
    For k = (8) To 0 Step -(1)
      mRange = k ^ 1.5
        GoSub Mosaic
    Next k
End Select
mosaicgoing = False
Exit Sub
'Actual Mosaic Code
'======================================================================
Mosaic:
'Get the bitmap info again, in case something's changed
GetObjectAPI pctMosaic.Picture, Len(bmp), bmp
'Reinit the SA
With SA
  .cbElements = 1
  .cDims = 2
  .bounds(0).lLbound = 0
  .bounds(0).cElements = bmp.bmHeight
  .bounds(1).lLbound = 0
  .bounds(1).cElements = bmp.bmWidthBytes
  .pvData = bmp.bmBits
End With
''Fake' the pointer
CopyMemory ByVal VarPtrArray(pict), VarPtr(SA), 4
'Work out the distance between the square division grid, and the pixel to get data from.
pCenter = (mRange) \ 2
'Find the limits of the image
uC = UBound(pict, 1)
uR = UBound(pict, 2)
For c = 0 To UBound(pict, 1) Step (mRange + 1)
  For r = 0 To UBound(pict, 2) Step (mRange + 1)
	  'Work out the distance between the square division grid, and the pixel to get data from.
      pCenter = (mRange) \ 2
	  'Pixel size to copy over
	  rRangei = (mRange)
      rRangej = (mRange)
      'Check if it's running out of bound, in case you turned the compiler option off.
      If c + mRange > UBound(pict, 1) Then rRangei = UBound(pict, 1) - c
      If r + mRange > UBound(pict, 2) Then rRangej = UBound(pict, 2) - r
      'Work out where to get the data from
      pC = c + pCenter
      pR = r + pCenter
      If pC > UBound(pict, 1) Then pC = c
      If pR > UBound(pict, 2) Then pR = r
      'Get the palette entry
      Value = PictureArray(pC, pR)
      If c = 0 Then cLimit = -pCenter
      If r = 0 Then rLimit = -pCenter
      'Copy the palette entry number over the region's pixels
      For i = cLimit To (rRangei)
        For j = rLimit To (rRangej)
          If c + i < 0 Or r + j < 0 Then GoTo SkipPixel
          pict(c + i, r + j) = Value
SkipPixel:
        Next j
      Next i
SkipThis:
  Next r
Next c
EndThis:
'Clean up
CopyMemory ByVal VarPtrArray(pict), 0&, 4
'Refresh, so the user sees the change. Don't replace with a DoEvents!
'Refreshing is slower, but it's less dangerous!
pctMosaic.Refresh
'======================================================================
Return
End Sub
```

