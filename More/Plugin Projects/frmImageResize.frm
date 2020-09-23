VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImageResize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Resize"
   ClientHeight    =   7305
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8460
   Icon            =   "frmImageResize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   564
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   6840
      Width           =   615
   End
   Begin ResizeImagePlugin.winConnect wc 
      Left            =   4440
      Top             =   3360
      _ExtentX        =   1535
      _ExtentY        =   661
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4440
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Images (*.bmp, *.gif, *jpg, *.ico, *.cur)|*.bmp;*.gif;*jpg;*.ico;*.cur"
      Flags           =   7
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   7320
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdResize 
      Caption         =   "Resize 16"
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00808080&
      Height          =   7080
      Left            =   120
      ScaleHeight     =   468
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   468
      TabIndex        =   2
      Top             =   120
      Width           =   7080
      Begin VB.PictureBox picSel 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         Height          =   1125
         Left            =   1080
         ScaleHeight     =   75
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   99
         TabIndex        =   10
         Top             =   3480
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   6760
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6760
         Width           =   255
      End
      Begin VB.HScrollBar hs 
         Height          =   260
         LargeChange     =   32
         Left            =   0
         Max             =   0
         SmallChange     =   16
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   6760
         Width           =   6760
      End
      Begin VB.VScrollBar vs 
         Height          =   6760
         LargeChange     =   32
         Left            =   6760
         Max             =   0
         SmallChange     =   16
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picMain 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         DrawStyle       =   2  'Dot
         Height          =   3285
         Left            =   120
         ScaleHeight     =   219
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   243
         TabIndex        =   3
         Top             =   120
         Width           =   3645
      End
   End
   Begin VB.CommandButton cmdResize 
      Caption         =   "Resize 32"
      Height          =   375
      Index           =   0
      Left            =   7320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   7320
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "Open Image"
   End
End
Attribute VB_Name = "frmImageResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Type Pixel
    Red As Double
    Green As Double
    Blue As Double
End Type

Dim mTag As String, mX1 As Integer, mY1 As Integer, mX2 As Integer, mY2 As Integer

Public Sub SmoothBlt(hdcSrc As Long, SrcX As Integer, SrcY As Integer, SrcWidth As Integer, SrcHeight As Integer, _
hdcDest As Long, DestX As Integer, DestY As Integer, DestWidth As Integer, DestHeight As Integer)
    Dim I As Double, J As Double
    Dim XFactor As Double, YFactor As Double
    
    'These factors are the ratio between the size
    'of the source and destination areas on each axis
    XFactor = SrcWidth / DestWidth
    YFactor = SrcHeight / DestHeight
    
    'From here its a simple double loop
    For I = 0 To DestWidth - 1
        For J = 0 To DestHeight - 1
            'the real legwork is in the GetFloatPix routine
            SetPixel hdcDest, I + DestX, J + DestY, GetFloatPix(hdcSrc, I * XFactor + SrcX, J * YFactor + SrcY)
        Next
        DoEvents
    Next
End Sub

'GetFloatPix is like GetPixel, except that it
'takes floating-point coordinates. This means
'that you can get color information approximated
'between pixels, allowing for smooth resizing.
Private Function GetFloatPix(hdc As Long, X As Double, Y As Double) As Long
    Dim P1 As Pixel, P2 As Pixel, P3 As Pixel, P4 As Pixel
    Dim Red As Long, Green As Long, Blue As Long
    Dim A1 As Double, A2 As Double, A3 As Double, A4 As Double
    Dim XPart As Double, YPart As Double
    
    'This handles the easy case of the pixel
    'being right on a spot.
    If X = Int(X) And Y = Int(Y) Then
        GetFloatPix = GetPixel(hdc, X, Y)
        Exit Function
    End If
    
    'This is the distance from the requested
    'point to the closest pixel to the upper-
    'left.
    XPart = X - Int(X)
    YPart = Y - Int(Y)
        
    'Any point the doesn't fall on a pixel will have
    'four surrounding pixels. We get their color data
    'here, in convenient Pixel structures.
    P1 = BreakPix(GetPixel(hdc, Int(X), Int(Y)))
    P2 = BreakPix(GetPixel(hdc, Int(X) + 1, Int(Y)))
    P3 = BreakPix(GetPixel(hdc, Int(X), Int(Y + 0.999)))
    P4 = BreakPix(GetPixel(hdc, Int(X + 0.999), Int(Y + 0.999)))
    
    'If we image the four pixels as adjoining squares, and
    'the pixel we want as a square overlapping them, there
    'are rectangles formed by the intersection of these
    'squares. We get the area of each rectangle here.
    A1 = (1 - XPart) * (1 - YPart)
    A2 = XPart * (1 - YPart)
    A3 = (1 - XPart) * YPart
    A4 = XPart * YPart
    
    'The areas serve to scale each of the four component
    'pixels to form our desired floating-point pixel. We
    'work on each channel independantly to prevent overflow
    'contaminating the other channels.
    Red = P1.Red * A1 + P2.Red * A2 + P3.Red * A3 + P4.Red * A4
    Green = P1.Green * A1 + P2.Green * A2 + P3.Green * A3 + P4.Green * A4
    Blue = P1.Blue * A1 + P2.Blue * A2 + P3.Blue * A3 + P4.Blue * A4

    'This simply reconstitutes the channels into a return value
    GetFloatPix = Red + Green * &H100& + Blue * &H10000
End Function

Private Function BreakPix(ByVal CVal As Long) As Pixel
'returns rgb values
  BreakPix.Blue = Int(CVal / 65536)
  BreakPix.Green = Int((CVal - (65536 * BreakPix.Blue)) / 256)
  BreakPix.Red = CVal - (65536 * BreakPix.Blue + 256 * BreakPix.Green)
End Function

Private Sub cmdOk_Click()
If picIcon(0).Visible = True Then
 Set picIcon(0).Picture = picIcon(0).Image
 Call SavePicture(picIcon(0).Picture, mTag)
 Call wc.Send("<")
ElseIf picIcon(1).Visible = True Then
 Set picIcon(1).Picture = picIcon(1).Image
 Call SavePicture(picIcon(0).Picture, mTag)
 Call wc.Send(">")
End If
End

End Sub

Private Sub cmdResize_Click(Index As Integer)
Set picIcon(Index).Picture = LoadPicture()
picMain.Cls
If mX2 <> 0 And mY2 <> 0 Then
 Set picSel.Picture = LoadPicture()
 picSel.Width = mX2
 picSel.Height = mY2
 
 Call picSel.PaintPicture(picMain.Picture, 0, 0, mX2, mY2, mX1, mY1, mX2, mY2, vbSrcCopy)
 picSel.Refresh
 Set picSel.Picture = picSel.Image
 
 Call SmoothBlt(picSel.hdc, 0, 0, picSel.ScaleWidth, picSel.ScaleHeight, picIcon(Index).hdc, 0, 0, 32, 32)
 picIcon(1).Visible = False
 picIcon(0).Visible = False
 picIcon(Index).Visible = True
 picIcon(Index).Refresh

 picMain.Line (mX1, mY1)-(mX1 + mX2, mY1 + mY2), 0, B
 picMain.Refresh
Else
 Call SmoothBlt(picMain.hdc, 0, 0, picMain.ScaleWidth, picMain.ScaleHeight, picIcon(Index).hdc, 0, 0, 32, 32)
 picIcon(1).Visible = False
 picIcon(0).Visible = False
 picIcon(Index).Visible = True
 picIcon(Index).Refresh

 picMain.Line (mX1, mY1)-(mX1 + mX2, mY1 + mY2), 0, B
 picMain.Refresh
End If
End Sub

Private Sub Form_Load()
If Command = "" Then
 MsgBox "This plugin cannot be run as a stand alone program.", vbInformation, "Plugin"
 End
End If

Call wc.SetCommand(Command)
End Sub

Private Sub hs_Change()
picMain.Left = -hs.Value + 6
End Sub

Private Sub mnuOpen_Click()
On Error GoTo 1
cd.ShowOpen

Set picMain.Picture = LoadPicture(cd.FileName)
Set picMain.Picture = picMain.Image
1
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMain.Cls
mX2 = 0
mY2 = 0
mX1 = X
mY1 = Y
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then Exit Sub
picMain.Cls
picMain.Line (mX1, mY1)-(X, Y), 0, B
Caption = "Image Resize [" & mX1 & ", " & mY1 & " > " & (X - mX1) & ", " & (Y - mY1) & "]"
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then Exit Sub
mX2 = X - mX1
mY2 = Y - mY1
End Sub

Private Sub picMain_Resize()
hs.Value = 0
hs.Max = picMain.Width - 237
vs.Value = 0
vs.Max = picMain.Height - 237
End Sub

Private Sub vs_Change()
picMain.Top = -vs.Value + 6
End Sub

Private Sub wc_Got(ByVal Msg As String)
Select Case Left(Msg, 1)
 Case "$"
  mTag = Mid(Msg, 2)
  Me.Show
End Select
End Sub
