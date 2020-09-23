VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDropShadow 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Drop Shadow Plugin"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2970
   Icon            =   "frmDropShadow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   2400
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.HScrollBar hsShades 
      Height          =   255
      Index           =   1
      LargeChange     =   25
      Left            =   120
      Max             =   255
      TabIndex        =   6
      Top             =   2880
      Value           =   1
      Width           =   1575
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
      Index           =   0
      LargeChange     =   25
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   2280
      Value           =   1
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
   Begin DropShadow.winConnect wc 
      Left            =   600
      Top             =   1440
      _ExtentX        =   1535
      _ExtentY        =   661
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y Offset: 1"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   750
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X Offset: 1"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   750
   End
   Begin VB.Image imgPreview 
      Height          =   1920
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "frmDropShadow"
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
Dim arrClr(1, 32, 32) As Long

For Y = 0 To picMain.ScaleHeight - 1
 For X = 0 To picMain.ScaleWidth - 1
    If picDefault.Point(X, Y) <> 8420352 Then arrClr(1, X, Y) = Picture1.BackColor Else arrClr(1, X, Y) = 8420352
    arrClr(0, X, Y) = picDefault.Point(X, Y)
 Next X
Next Y

For Y = 0 To picMain.Height
 For X = 0 To picMain.Width
  If X > hsShades(0).Value And Y > hsShades(1).Value Then
   picMain.PSet (X, Y), arrClr(1, X - hsShades(0).Value, Y - hsShades(1).Value)
  End If
  If arrClr(0, X, Y) <> 8420352 Then picMain.PSet (X, Y), arrClr(0, X, Y)
 Next X
Next Y

picMain.Refresh
Set picMain.Picture = picMain.Image
Set imgPreview.Picture = picMain.Picture
End Sub

Private Sub Command2_Click()
Call SavePicture(imgPreview.Picture, picDefault.Tag)
Call wc.Send("!")
End
End Sub

Private Sub Form_Load()
If Command = "" Then
 MsgBox "This plugin cannot be run as a stand alone program.", vbInformation, "Plugin"
 End
End If

Call wc.SetCommand(Command)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hsShades_Change(Index As Integer)
Dim arr(1) As String
arr(0) = "X"
arr(1) = "Y"
lbl(Index).Caption = arr(Index) & " Offset: " & hsShades(Index).Value
End Sub

Private Sub Picture1_Click()
CD.ShowColor
Picture1.BackColor = CD.Color
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
