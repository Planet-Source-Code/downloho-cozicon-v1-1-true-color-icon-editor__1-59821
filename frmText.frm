VERSION 5.00
Begin VB.Form frmText 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text Tool"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4800
      MouseIcon       =   "frmText.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   960
      Width           =   480
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1455
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   2175
      Begin VB.ComboBox cmbSize 
         Height          =   315
         ItemData        =   "frmText.frx":0CCA
         Left            =   840
         List            =   "frmText.frx":0CFE
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox cmbFonts 
         Height          =   315
         ItemData        =   "frmText.frx":0D40
         Left            =   600
         List            =   "frmText.frx":0D42
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbY 
         Height          =   315
         ItemData        =   "frmText.frx":0D44
         Left            =   1440
         List            =   "frmText.frx":0DA5
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbX 
         Height          =   315
         ItemData        =   "frmText.frx":0E1C
         Left            =   360
         List            =   "frmText.frx":0E7D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Fontsize : "
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Font : "
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Y : "
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "X : "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.TextBox txtText 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbFonts_Click()
If cmbFonts.Text = "" Then Exit Sub
txtText.Font = cmbFonts.Text
Call rSetFont(cmbX.Text, cmbY.Text, cmbFonts.Text, cmbSize.Text, txtText.Text)
End Sub

Private Sub cmbSize_Click()
txtText.FontSize = cmbSize.Text
Call rSetFont(cmbX.Text, cmbY.Text, cmbFonts.Text, cmbSize.Text, txtText.Text)
End Sub

Private Sub cmbX_Click()
On Error Resume Next
Call rSetFont(cmbX.Text, cmbY.Text, cmbFonts.Text, cmbSize.Text, txtText.Text)
End Sub

Private Sub cmbY_Click()
On Error Resume Next
Call rSetFont(cmbX.Text, cmbY.Text, cmbFonts.Text, cmbSize.Text, txtText.Text)
End Sub

Private Sub cmdCan_Click()
Me.Hide
Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
Call frmMain.SetFont(cmbX.Text, cmbY.Text, cmbFonts.Text, cmbSize.Text, txtText.Text)
Me.Hide
Call Unload(Me)
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 1 To Screen.FontCount
 cmbFonts.AddItem Screen.Fonts(i)
Next i
cmbFonts.ListIndex = 0
cmbX.ListIndex = 0
cmbY.ListIndex = 0
cmbSize.ListIndex = 0
End Sub

Private Sub rSetFont(ByVal x As Integer, ByVal y As Integer, ByVal FontName As String, ByVal FontSize As Integer, ByVal Text As String)
If FontName = "" Then Exit Sub
picTemp.Cls
With picTemp
 .CurrentX = x - 1
 .CurrentY = y - 4
 .ForeColor = frmMain.picClr(1).BackColor
 .Font = FontName
 .FontSize = FontSize
End With
 
  Dim arr() As String, v As Variant
  arr() = Split(Text, vbCrLf)
  For Each v In arr()
   picTemp.Print v
   Debug.Print FontSize
   picTemp.CurrentY = picTemp.CurrentY + (FontSize / 5) - 5
   picTemp.CurrentX = x - 1
  Next v
End Sub

Private Sub txtText_Change()
Call rSetFont(cmbX.Text, cmbY.Text, cmbFonts.Text, cmbSize.Text, txtText.Text)
End Sub
