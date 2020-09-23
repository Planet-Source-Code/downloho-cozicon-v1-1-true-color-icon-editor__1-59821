VERSION 5.00
Begin VB.Form frmGrayScale 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gray Scale Plugin"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2970
   Icon            =   "frmGrayScale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin GrayScale.winConnect wc 
      Left            =   2280
      Top             =   1320
      _ExtentX        =   1535
      _ExtentY        =   661
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picDefault 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.HScrollBar hsShades 
      Height          =   255
      LargeChange     =   25
      Left            =   120
      Max             =   255
      Min             =   2
      TabIndex        =   2
      Top             =   2280
      Value           =   16
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shades of Gray: 16"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Image imgPreview 
      Height          =   1920
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "frmGrayScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type

Dim clrs() As Long

Private Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function

Private Function GetClosetClr(ByVal Clr As Long) As Integer
Dim c&, d&, e&, s&, i%, l&, u%

 c& = GetRGB(Clr&).Red: d& = GetRGB(Clr&).Green: e& = GetRGB(Clr&).Blue

 s& = 30000

For i% = 0 To UBound(clrs())
  l& = Abs(clrs(i%, 0) - c) + Abs(clrs(i%, 1) - d) + Abs(clrs(i%, 2) - e)
  If l& < s& Then s& = l&: u% = i%
Next i%
GetClosetClr% = u%
End Function

Private Sub Command1_Click()
ReDim clrs(hsShades.Value - 1, 2) As Long
For i = 0 To hsShades.Value - 1
 d = Int((255 * i) / (hsShades.Value - 1))
 clrs(i, 0) = d
 clrs(i, 1) = d
 clrs(i, 2) = d
Next

For y = 0 To picMain.Height - 1
 For x = 0 To picMain.Width - 1
  u = GetClosetClr(picDefault.Point(x, y))
  Caption = picDefault.Point(x, y)
  If picDefault.Point(x, y) <> CLng(8421376) Then picMain.PSet (x, y), RGB(clrs(u, 0), clrs(u, 1), clrs(u, 2))
 Next x
Next y
picMain.Refresh
Set picMain.Picture = picMain.Image
Set imgPreview.Picture = picMain.Picture
'Call SavePicture(picMain.Picture, "c:\windows\desktop\temp.bmp")
End Sub

Private Sub Command2_Click()
Call SavePicture(imgPreview.Picture, picDefault.Tag)
Call wc.Send("!")
End
End Sub

Private Sub Form_Load()
Call wc.SetCommand(Command)
ReDim clrs(15, 2) As Long
For i = 0 To 15
 d = Int((255 * i) / 15)
 clrs(i, 0) = d
 clrs(i, 1) = d
 clrs(i, 2) = d
Next
If Command = "" Then
  Set picMain.Picture = LoadPicture("c:\windows\desktop\temp.bmp")
  Set picDefault.Picture = picMain.Picture
  Set imgPreview.Picture = picMain.Picture
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hsShades_Change()
lbl.Caption = "Shades of Gray: " & hsShades.Value
End Sub

Private Sub wc_Got(ByVal Msg As String)
Select Case Left(Msg, 1)
 Case "$"
  picDefault.Tag = Mid(Msg, 2)
  Set picMain.Picture = LoadPicture(Mid(Msg, 2))
  Set picDefault.Picture = picMain.Picture
  Set imgPreview.Picture = picMain.Picture
  Me.Show
End Select
End Sub
