VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColorReplace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Replace Plugin"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2970
   Icon            =   "frmColorReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ColorReplacePlugin.winConnect wc 
      Left            =   1200
      Top             =   2760
      _ExtentX        =   1535
      _ExtentY        =   661
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      Height          =   1980
      Left            =   120
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   5
      Top             =   120
      Width           =   1980
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2400
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   7
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   975
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
      Caption         =   "Replace Color:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1050
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color: "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   450
   End
End
Attribute VB_Name = "frmColorReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub Command1_Click()
For y = 0 To picMain.ScaleHeight - 1
 For X = 0 To picMain.ScaleWidth - 1
    If picDefault.Point(X, y) = picClr(0).BackColor Then
     picMain.PSet (X, y), picClr(1).BackColor
    End If
 Next X
Next y

  picMain.Refresh
  Set picMain.Picture = picMain.Image
  Call StretchBlt(picPreview.hdc, 0, 0, 128, 128, picMain.hdc, 0, 0, 32, 32, vbSrcCopy)
  picPreview.Refresh

End Sub

Private Sub Command2_Click()
Set picMain.Picture = picMain.Image
Call SavePicture(picMain.Picture, picDefault.Tag)
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

Private Sub picClr_Click(Index As Integer)
On Error GoTo 1
CD.ShowColor
picClr(Index).BackColor = CD.Color

For y = 0 To picMain.ScaleHeight - 1
 For X = 0 To picMain.ScaleWidth - 1
    If picDefault.Point(X, y) = picClr(0).BackColor Then
     picMain.PSet (X, y), picClr(1).BackColor
    End If
 Next X
Next y

  picMain.Refresh
  Set picMain.Picture = picMain.Image
  Call StretchBlt(picPreview.hdc, 0, 0, 128, 128, picMain.hdc, 0, 0, 32, 32, vbSrcCopy)
  picPreview.Refresh

1
End Sub

Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then picClr(0).BackColor = picPreview.Point(X, y)
End Sub

Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then picClr(0).BackColor = picPreview.Point(X, y)
End Sub

Private Sub wc_Got(ByVal Msg As String)
Select Case Left(Msg, 1)
 Case "$"
  picDefault.Tag = Mid(Msg, 2)
  Set picMain.Picture = LoadPicture(Mid(Msg, 2))
  Set picDefault.Picture = picMain.Picture
  Call StretchBlt(picPreview.hdc, 0, 0, 128, 128, picDefault.hdc, 0, 0, 32, 32, vbSrcCopy)
  picPreview.Refresh
  Me.Show
End Select
End Sub
