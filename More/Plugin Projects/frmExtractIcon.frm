VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExtractIcon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extract Icon Plugin"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2970
   Icon            =   "frmExtractIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   300
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2400
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Executables (*.exe, *.dll)|*.exe;*.dll"
      Flags           =   7
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin ExtractIcon.winConnect wc 
      Left            =   600
      Top             =   1440
      _ExtentX        =   1535
      _ExtentY        =   661
   End
   Begin VB.Image imgPreview 
      Height          =   1920
      Left            =   120
      Picture         =   "frmExtractIcon.frx":0CCC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "frmExtractIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

Private Type typSHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
    End Type

Private Enum iconsize
    largeIcon = 0
    SmallIcon = 1
End Enum

Private Const SH_USEFILEATTRIBUTES As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const SHGFI_DISPLAYNAME As Long = &H200
Private Const SHGFI_EXETYPE As Long = &H2000
Private Const SHGFI_SYSICONINDEX As Long = &H4000
Private Const SHGFI_SHELLICONSIZE As Long = &H4
Private Const SHGFI_TYPENAME As Long = &H400
Private Const SHGFI_LARGEICON As Long = &H0
Private Const SHGFI_SMALLICON As Long = &H1
Private Const ILD_TRANSPARENT As Long = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE Or SH_USEFILEATTRIBUTES
Private FileInfo As typSHFILEINFO

Private Function geticonhandle(filename As String, size As Long)  'Gets a handle To the icon
    geticonhandle = SHGetFileInfo(filename, FILE_ATTRIBUTE_NORMAL, FileInfo, Len(FileInfo), Flags Or size)
End Function

Private Function drawfileicon(filetype As String, size As iconsize, destHDC As Long, x As Long, y As Long)  'Draws the icon int the destination.hdc
    drawfileicon = ImageList_Draw(geticonhandle(filetype, size), FileInfo.iIcon, destHDC, x, y, ILD_TRANSPARENT)
End Function

Private Sub Command1_Click()
'On Error GoTo 1

CD.ShowOpen
Set picMain.Picture = LoadPicture()
Call drawfileicon(CD.filename, largeIcon, picMain.hDC, 0, 0)
Text1.Text = CD.filename
picMain.Refresh
Set picMain.Picture = picMain.Image
Set imgPreview.Picture = picMain.Picture
1
End Sub

Private Sub Command2_Click()
Call SavePicture(imgPreview.Picture, picMain.Tag)
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

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub wc_Got(ByVal Msg As String)
Select Case Left(Msg, 1)
 Case "$"
  picMain.Tag = Mid(Msg, 2)
  Me.Show
End Select
End Sub
