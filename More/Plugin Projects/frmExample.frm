VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSend 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   600
      Picture         =   "frmExample.frx":0D4A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   4560
      Width           =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send to CozIcon"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New 16x16 Icon in CozIcon"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New 32x32 Icon in CozIcon"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
   Begin VB.CommandButton cmdGetAll 
      Caption         =   "Get All Icons"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   2175
   End
   Begin ExamplePlugin.winConnect wc 
      Left            =   1440
      Top             =   840
      _ExtentX        =   1535
      _ExtentY        =   661
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Icon to Send: "
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "All Icons:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Default Icon: "
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSaveFileAs As String

Private Sub cmdGetAll_Click()
'send command to get all icons
wc.Send "+"
End Sub

Private Sub Command1_Click()
'request a new 32x32 icon
wc.Send "<"
End Sub

Private Sub Command2_Click()
'request a new 16x16 icon
wc.Send ">"
End Sub

Private Sub Command3_Click()
'send image to CozIcon
Call SavePicture(picSend.Picture, mSaveFileAs) 'save the picture
wc.Send "!" 'let CozIcon know to load it
End Sub

Private Sub Form_Load()
If Command = "" Then 'end the program if ran as a stand alone app (optional)
 MsgBox "This plugin cannot be run as a stand alone program.", vbInformation, "Plugin"
 End
End If
Call wc.SetCommand(Command) 'set command
'basically CozIcon sends a Message containing the hwnd of WinConnect.
'The WinConnect in this program parse it and sends it's hwnd.
'CozIcon then sends the current icon. (see: Sub wc_Got)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub wc_Got(ByVal Msg As String)
Select Case Left(Msg, 1)
 Case "$" 'default icon
  mSaveFileAs = Mid(Msg, 2) 'file that CozIcon checks to load
  Set pic(0).Picture = LoadPicture(mSaveFileAs) 'it also saves the current icon to this file
  Me.Show 'show form
 Case "-" 'all icons in CozIcon
   Dim i As Integer
   i = Mid(Msg, 2, InStr(3, Msg, "-") - 2)
   Call Load(pic(i))
   pic(i).Left = i * pic(i).ScaleWidth
   pic(i).Top = 88
   pic(i).Visible = True
   Set pic(i).Picture = LoadPicture(Mid(Msg, InStr(3, Msg, "-") + 1))
End Select
End Sub

