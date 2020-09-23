Attribute VB_Name = "modFileIcon"
Option Explicit

Public Type JPALETTE
 jpRed As Long
 jpGreen As Long
 jpBlue As Long
 jpColor As Long
End Type

Public Type ICONDATAPREFIX
 idpLength As Long
 idpWidth As Integer
 idpHeightT2 As Integer
 idpColorDepth As Integer
End Type

Public Type ICONDATA
 idWidth As Integer
 idHeight As Integer
 idColorCount As Integer
 idColorCount2 As Long
 idPalette() As JPALETTE
 idDataLength As Long
 idDataOffset As Long
 idPrefix As ICONDATAPREFIX
 idData As String
End Type

Public Type ICONFILEDATA
 ifdCount As Integer
 ifdIconData() As ICONDATA
 ifdIcon() As Picture
 ifdSuccess As Boolean
End Type

Public Temp_Icon As ICONFILEDATA

Public Function GenerateIconForSave(ByRef PicBox As PictureBox) As String
'method for generating the binary data required for TRUE COLOR(!) icons

Dim arrX() As String, isTr As Boolean, bpp As Integer
Dim X As Integer, Y As Integer, s As String
Dim width_in_pixels As Integer, height_in_pixels As Integer
Dim l As Long, r As Integer, g As Integer, b As Integer

PicBox.ScaleMode = vbPixels
width_in_pixels = PicBox.ScaleWidth
height_in_pixels = PicBox.ScaleHeight
bpp = GetBinaryBitCount(width_in_pixels) - 1
ReDim arrX(height_in_pixels - 1, bpp) As String

For Y = height_in_pixels - 1 To 0 Step -1 'icons are saved from top to bottom
 For X = 0 To width_in_pixels - 1         'and left to right
  l = PicBox.Point(X, Y)
  
  r = GetRGB(l).Red
  g = GetRGB(l).Green
  b = GetRGB(l).Blue
  
  If r < 0 Then r = 0
  If g < 0 Then g = 0
  If b < 0 Then b = 0
  
  If l = PicBox.BackColor Then
   s = s & Chr(0) & Chr(0) & Chr(0)
   isTr = True
  Else
   s = s & Chr(b) & Chr(g) & Chr(r)
  End If
 Next X
Next Y

If isTr = True Then 'transparent
 For Y = height_in_pixels - 1 To 0 Step -1 'populate transparent data array
  For X = 0 To bpp                           'to make sure that'll fill properly for 16x16 icons
   arrX(Y, X) = "-1"
  Next X
 Next Y

For Y = 0 To height_in_pixels - 1 'check for transparency
 For X = 0 To width_in_pixels - 1
  l = PicBox.Point(X, Y)
  If l = PicBox.BackColor Then 'generate transparent string for transsolution function
    If arrX(Y, Int(X / 8)) = "-1" Then arrX(Y, Int(X / 8)) = ""
    arrX(Y, Int(X / 8)) = arrX(Y, Int(X / 8)) & "1"
  Else
    If arrX(Y, Int(X / 8)) = "-1" Then arrX(Y, Int(X / 8)) = ""
    arrX(Y, Int(X / 8)) = arrX(Y, Int(X / 8)) & "0"
  End If
  
 Next X
Next Y

Dim f As String, e As Integer
 For Y = height_in_pixels - 1 To 0 Step -1 'create generated transparent data
  For X = 0 To bpp
   If arrX(Y, X) = "-1" Then e = 255 Else e = BinToDec(arrX(Y, X))
   f = f & Chr(e)
  Next X
 Next Y
End If

 s = s & f
 'fill data with chr(0) if there is no transparency
 s = s & String(((width_in_pixels * height_in_pixels) * 3) + ((width_in_pixels / 4) * width_in_pixels) - Len(s), Chr(0))
 Dim icon_id As String
 Dim icon_count As String
 Dim icon_header As String
 Dim icon_position As String
 Dim icon_before_data As Integer
 
 icon_id = String(2, Chr(0)) & Chr(1) & Chr(0)
 icon_count = Chr(1) & Chr(0)
 'icon_header
 icon_position = Chr(22) & Chr(0) & Chr(0) & Chr(0)
 icon_before_data = 40
If width_in_pixels = 16 Then '16x16 icon
 'icon_header = Chr(width_in_pixels) & Chr(height_in_pixels) & String(6, Chr(0)) & Chr(Val("&H68")) & Chr(3) & String(2, Chr(0))
 icon_header = Chr(width_in_pixels) & Chr(height_in_pixels) & String(6, Chr(0)) & Long2Chr(icon_before_data + Len(s)) & String(2, Chr(0)) & icon_position
 GenerateIconForSave = icon_id & icon_count & icon_header & Chr(icon_before_data) & String(3, Chr(0)) & Chr(16) & String(3, Chr(0)) & Chr(32) & String(3, Chr(0)) & Chr(1) & Chr(0) & Chr(24) & String(5, Chr(0)) & Chr(64) & Chr(3) & String(18, Chr(0)) & s
ElseIf width_in_pixels = 32 Then '32x32 icon
 icon_header = Chr(width_in_pixels) & Chr(height_in_pixels) & String(6, Chr(0)) & Long2Chr(icon_before_data + Len(s)) & String(2, Chr(0)) & icon_position
 GenerateIconForSave = icon_id & icon_count & icon_header & Chr(icon_before_data) & String(3, Chr(0)) & Chr(32) & String(3, Chr(0)) & Chr(64) & String(3, Chr(0)) & Chr(1) & Chr(0) & Chr(24) & String(6, Chr(0)) & Chr(12) & String(18, Chr(0)) & s
ElseIf width_in_pixels = 48 Then '48x48 icon
 icon_header = Chr(width_in_pixels) & Chr(height_in_pixels) & String(6, Chr(0)) & Long2Chr(icon_before_data + Len(s)) & String(2, Chr(0)) & icon_position
 GenerateIconForSave = icon_id & icon_count & icon_header & Chr(icon_before_data) & String(3, Chr(0)) & Chr(48) & String(3, Chr(0)) & Chr(96) & String(3, Chr(0)) & Chr(1) & Chr(0) & Chr(24) & String(6, Chr(0)) & Chr(12) & String(18, Chr(0)) & s
Else 'unsupported icon size
 Err.Raise 10001, , "Unsupported Icon Size"
End If
End Function

Public Function GenerateIconForSaveX(ByRef PicBox As PictureBox) As String
'method for generating the binary data required for TRUE COLOR(!) icons

Dim arrX() As String, isTr As Boolean
Dim X As Integer, Y As Integer, s As String
Dim width_in_pixels As Integer, height_in_pixels As Integer
Dim l As Long, r As Integer, g As Integer, b As Integer

PicBox.ScaleMode = vbPixels
width_in_pixels = PicBox.ScaleWidth
height_in_pixels = PicBox.ScaleHeight

Dim bpp As Integer
bpp = GetBinaryBitCount(width_in_pixels) - 1
ReDim arrX(height_in_pixels - 1, bpp) As String

For Y = height_in_pixels - 1 To 0 Step -1 'icons are saved from top to bottom
 For X = 0 To width_in_pixels - 1         'and left to right
  l = PicBox.Point(X, Y)
  
  r = GetRGB(l).Red
  g = GetRGB(l).Green
  b = GetRGB(l).Blue
  
  If r < 0 Then r = 0
  If g < 0 Then g = 0
  If b < 0 Then b = 0
  
  If l = PicBox.BackColor Then
   s = s & Chr(0) & Chr(0) & Chr(0)
   isTr = True
  Else
   s = s & Chr(b) & Chr(g) & Chr(r)
  End If
 Next X
Next Y

If isTr = True Then 'transparent
 For Y = height_in_pixels - 1 To 0 Step -1 'populate transparent data array
  For X = 0 To 3                           'to make sure that'll fill properly for 16x16 icons
   arrX(Y, X) = "-1"
  Next X
 Next Y

For Y = 0 To height_in_pixels - 1 'check for transparency
 For X = 0 To width_in_pixels - 1
  l = PicBox.Point(X, Y)
  If l = PicBox.BackColor Then 'generate transparent string for transsolution function
    If arrX(Y, Int(X / 8)) = "-1" Then arrX(Y, Int(X / 8)) = ""
    arrX(Y, Int(X / 8)) = arrX(Y, Int(X / 8)) & "1"
  Else
    If arrX(Y, Int(X / 8)) = "-1" Then arrX(Y, Int(X / 8)) = ""
    arrX(Y, Int(X / 8)) = arrX(Y, Int(X / 8)) & "0"
  End If
  
 Next X
Next Y

Dim f As String, e As Integer
 For Y = height_in_pixels - 1 To 0 Step -1 'create generated transparent data
  For X = 0 To 3
   If arrX(Y, X) = "-1" Then e = 255 Else e = BinToDec(arrX(Y, X))
   f = f & Chr(e)
  Next X
 Next Y
End If

 s = s & f
 'fill data with chr(0) if there is no transparency
 s = s & String(((width_in_pixels * height_in_pixels) * 3) + ((width_in_pixels / 4) * width_in_pixels) - Len(s), Chr(0))
 Dim icon_before_data As Integer
 icon_before_data = 40
If width_in_pixels = 16 Then '16x16 icon
 GenerateIconForSaveX = Chr(icon_before_data) & String(3, Chr(0)) & Chr(16) & String(3, Chr(0)) & Chr(32) & String(3, Chr(0)) & Chr(1) & Chr(0) & Chr(24) & String(5, Chr(0)) & Chr(64) & Chr(3) & String(18, Chr(0)) & s
ElseIf width_in_pixels = 32 Then '32x32 icon
 GenerateIconForSaveX = Chr(icon_before_data) & String(3, Chr(0)) & Chr(32) & String(3, Chr(0)) & Chr(64) & String(3, Chr(0)) & Chr(1) & Chr(0) & Chr(24) & String(6, Chr(0)) & Chr(12) & String(18, Chr(0)) & s
Else 'unsupported icon size
 Err.Raise 10001, , "Unsupported Icon Size"
End If
End Function

Public Function GetIconSize(ByVal Filename As String) As Integer
'check if a file is an icon AND if it is return the icon size
Dim s As String, l As Long
l = FreeFile()
Open Filename For Binary Access Read As #l
 s = Input(LOF(l), #l)
Close #l
If Left(s, 4) <> Chr(0) & Chr(0) & Chr(1) & Chr(0) Then 'not an icon
 GetIconSize = -1
Else 'is an icon get the size
 GetIconSize = Asc(Mid(s, 7, 1))
End If
End Function

Private Function RePow(i As Integer, k As Integer) As Integer
Dim j As Integer, Count As Integer
  If k > 0 Then j = 2 Else j = 1
  For Count = 1 To k - 1
    j = j * 2
  Next Count
  RePow = j
End Function

Private Function BinToDec(s As String) As Integer
'replaced TransSoloution, figure out it really was a binary string
Dim iL As Integer, a As Integer, i As Integer
  iL = Len(s)
  a = 0
  For i = 1 To iL
    If Mid(s, i, 1) = "0" Or Mid(s, i, 1) = "1" Then
      a = a + RePow(2, iL - i) * CInt(Mid(s, i, 1))
    Else
      Exit Function
    End If
  Next i
BinToDec = a
End Function

Private Function Chr2Long(ByVal s As String) As Long
Dim i As Integer, l As Long
For i = Len(s) To 1 Step -1
 If i = 1 Then
  l = l + Asc(Mid(s, 1, 1))
 Else
  l = l + (Asc(Mid(s, i, 1)) * 256)
 End If
Next i

Chr2Long = l
End Function

Public Function Long2Chr(ByVal Num As Long) As String
Dim lA As Long, lB As Long, lC As Long

 lB = Num
 lA = Int(lB / 256)
 lB = lB - (lA * 256)

 If lA > 255 Then
  lC = Int(lA / 256)
  lA = lA - (lC * 256)
 End If
 
 Long2Chr = Chr(lB) & IIf(lC <> 0, Chr(lC), "") & IIf(lA <> 0, Chr(lA), "")
End Function

Public Sub KillFile(ByVal Filename As String)
On Error GoTo 1
Call Kill(Filename)
1
End Sub

Private Function SaveIconSimple(dIconData As ICONDATA, ByVal Filename As String) As Picture
'saves icon and returns the picture from a valid ICONDATA variable
On Error GoTo 1
Dim s As String
Dim l As Long

s = Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(1) & Chr(0) & Chr(dIconData.idWidth) & Chr(dIconData.idHeight) & Chr(dIconData.idColorCount) & String(5, Chr(0)) & _
    Long2Chr(dIconData.idDataLength) & Chr(0) & Chr(0) & Chr(22) & Chr(0) & Chr(0) & Chr(0) & _
    dIconData.idData
Call KillFile(Filename)
l = FreeFile()
Open Filename For Binary Access Write As #l
 Put #l, , s
Close #l

Set SaveIconSimple = LoadPicture(Filename)
1
End Function

Public Function LoadIcon(ByVal Filename As String) As ICONFILEDATA
'for load icons
'checks for multiple icons inside a file and will extract them all.
Dim l As Long, a As Long
Dim s As String, f As String
Dim temp As ICONFILEDATA

l = FreeFile()
Open Filename For Binary Access Read As #l
 s = Input(LOF(l), #l)
Close #l

If Left(s, 4) <> Chr(0) & Chr(0) & Chr(1) & Chr(0) Then LoadIcon = temp: Exit Function

temp.ifdCount = Asc(Mid(s, 5, 1))

For l = 0 To temp.ifdCount - 1
 ReDim Preserve temp.ifdIconData(l) As ICONDATA
 ReDim Preserve temp.ifdIcon(l) As Picture

 f = Mid(s, 6 + (l * 16) + 1, 16)
 
 temp.ifdIconData(l).idWidth = Asc(Mid(f, 1, 1))
 temp.ifdIconData(l).idHeight = Asc(Mid(f, 2, 1))
 temp.ifdIconData(l).idColorCount = Asc(Mid(f, 3, 1))
 temp.ifdIconData(l).idDataLength = Chr2Long(Mid(f, 9, 3)) '- 1
 temp.ifdIconData(l).idDataOffset = Chr2Long(Mid(f, 13, 3)) + 1
 
 temp.ifdIconData(l).idData = Mid(s, temp.ifdIconData(l).idDataOffset, temp.ifdIconData(l).idDataLength + 1)
 Set temp.ifdIcon(l) = SaveIconSimple(temp.ifdIconData(l), App.Path & "\temp.ico")
 
 a = Asc(Mid(temp.ifdIconData(l).idData, 15, 1))
 
 Select Case a
  Case 4
   temp.ifdIconData(l).idColorCount2 = 16
   f = Mid(temp.ifdIconData(l).idData, 41, 64)
   For a = 0 To 15
    ReDim Preserve temp.ifdIconData(l).idPalette(a)
    temp.ifdIconData(l).idPalette(a).jpBlue = Asc(Mid(f, 1, 1))
    temp.ifdIconData(l).idPalette(a).jpGreen = Asc(Mid(f, 2, 1))
    temp.ifdIconData(l).idPalette(a).jpRed = Asc(Mid(f, 3, 1))
    temp.ifdIconData(l).idPalette(a).jpColor = RGB(Asc(Mid(f, 3, 1)), Asc(Mid(f, 2, 1)), Asc(Mid(f, 1, 1)))
    f = Mid(f, 5)
   Next a
  Case 8
   temp.ifdIconData(l).idColorCount2 = 256
   
   '41
   '1024
   f = Mid(temp.ifdIconData(l).idData, 41, 1024)
   For a = 0 To 255
    ReDim Preserve temp.ifdIconData(l).idPalette(a)
    temp.ifdIconData(l).idPalette(a).jpBlue = Asc(Mid(f, 1, 1))
    temp.ifdIconData(l).idPalette(a).jpGreen = Asc(Mid(f, 2, 1))
    temp.ifdIconData(l).idPalette(a).jpRed = Asc(Mid(f, 3, 1))
    temp.ifdIconData(l).idPalette(a).jpColor = RGB(Asc(Mid(f, 3, 1)), Asc(Mid(f, 2, 1)), Asc(Mid(f, 1, 1)))
    f = Mid(f, 5)
   Next a
  Case 24
   temp.ifdIconData(l).idColorCount2 = 16777216
 End Select

Next l
temp.ifdSuccess = True
LoadIcon = temp
End Function

Public Sub GeneratePalette(Pic As PictureBox, ByRef jpArr() As JPALETTE)
'transparent color is pictures backcolor
Dim c As Long
Dim Y As Integer, X As Integer
Dim i As Integer, b As Boolean
Dim k As Integer
'first color must be black
ReDim jpArr(k)
For Y = 0 To Pic.ScaleHeight - 1
 For X = 0 To Pic.ScaleWidth - 1
 
  c = Pic.Point(X, Y)
  If c <> Pic.BackColor Then
   b = False
   For i = 0 To UBound(jpArr())
    If c = jpArr(i).jpColor Then b = True: Exit For
   Next i

   If b = False Then
    k = k + 1
    ReDim Preserve jpArr(k)
    If k > 255 Then Exit Sub
    jpArr(k).jpColor = c
    jpArr(k).jpBlue = Int(c / 65536)
    jpArr(k).jpGreen = Int((c - (65536 * jpArr(k).jpBlue)) / 256)
    jpArr(k).jpRed = c - (65536 * jpArr(k).jpBlue + 256 * jpArr(k).jpGreen)
   End If
  End If
 Next X
Next Y
End Sub

Public Function SavePalette(ByRef jpArr() As JPALETTE, Optional ByVal Filename As String) As String
'expects 'a' to be an array filled with long colors
Dim s As String
If Filename <> "" Then
If UBound(jpArr) = 15 Then
 s = "RIFFP" & Chr(0) & String(3, Chr(0)) & "PAL dataD" & String(4, Chr(0)) & Chr(3) & Chr(16)
ElseIf UBound(jpArr) = 255 Then
 s = "RIFF" & Chr(16) & Chr(4) & String(2, Chr(0)) & "PAL data" & Chr(4) & Chr(4) & String(3, Chr(0)) & Chr(3) & Chr(0) & Chr(1)
Else
 Exit Function
End If
End If

Dim i As Integer
For i = 0 To UBound(jpArr)
  s = s & Chr(jpArr(i).jpBlue) & Chr(jpArr(i).jpGreen) & Chr(jpArr(i).jpRed) & Chr(0)
Next i

If Filename <> "" Then
On Error Resume Next
Call Kill(Filename)
Open Filename For Binary Access Write As #1
 Put #1, , s$
Close #1
Else
 SavePalette = s
End If
End Function

Public Function GetColorIndexFromPalette(ByRef jpArr() As JPALETTE, ByVal Clr As Long) As Integer
Dim i As Integer
For i = 0 To UBound(jpArr())
 If Clr = jpArr(i).jpColor Then GetColorIndexFromPalette = i: Exit Function
Next i
GetColorIndexFromPalette = -1
End Function

Public Sub FillPalette(ByRef jpArr() As JPALETTE)
Dim k As Integer, i As Integer
k = UBound(jpArr())
If k <= 15 Then
 For i = k + 1 To 15
   ReDim Preserve jpArr(i)
   jpArr(i).jpBlue = 255
   jpArr(i).jpGreen = 255
   jpArr(i).jpRed = 255
   jpArr(i).jpColor = vbWhite
 Next i
ElseIf k <= 255 Then
 For i = k + 1 To 255
   ReDim Preserve jpArr(i)
   jpArr(i).jpBlue = 255
   jpArr(i).jpGreen = 255
   jpArr(i).jpRed = 255
   jpArr(i).jpColor = vbWhite
 Next i
End If
End Sub

Public Function GetBinaryBitCount(ByVal IconWidth As Integer) As Integer
Select Case IconWidth
 Case 0 To 32
  GetBinaryBitCount = 4
 Case 33 To 48
  GetBinaryBitCount = 8
End Select
End Function

Public Function GenerateIconDataPrefix(d As ICONDATAPREFIX) As String
'                        Chr(40) & String(3, Chr(0)) & Chr(32) & String(3, Chr(0)) & Chr(64) & String(3, Chr(0)) & Chr(1) & Chr(0) & Chr(8) & String(6, Chr(0)) & Chr(4) & String(18, Chr(0))
Dim s As String
GenerateIconDataPrefix = Chr(d.idpLength) & String(3, Chr(0)) & Chr(d.idpWidth) & String(3, Chr(0)) & Chr(d.idpHeightT2) & String(3, Chr(0)) & Chr(1) & Chr(0) & Chr(d.idpColorDepth) & String(25, Chr(0))
End Function

Public Function GetClosestColor(ByRef jpArr() As JPALETTE, ByVal Clr As Long) As Integer
Dim r As Long, g As Long, b As Long
Dim s As Long, i As Integer, l As Long, u As Integer

  b = Abs(Int(Clr / 65536))
  g = Abs(Int((Clr - (65536 * b)) / 256))
  r = Abs(Clr - (65536 * b + 256 * g))

 s = 300000000

For i = 0 To UBound(jpArr())
  l = Abs(jpArr(i).jpRed - r) + Abs(jpArr(i).jpGreen - g) + Abs(jpArr(i).jpBlue - b)
  If l < s Then s = l: u = i
Next i
GetClosestColor = u
End Function

