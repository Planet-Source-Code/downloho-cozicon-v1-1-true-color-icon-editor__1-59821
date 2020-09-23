VERSION 5.00
Begin VB.Form frmMIcons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Multiple Icons"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.VScrollBar vs 
      Height          =   3840
      LargeChange     =   32
      Left            =   3270
      Max             =   0
      SmallChange     =   10
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00C0E0FF&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   208
      TabIndex        =   2
      Top             =   120
      Width           =   3180
      Begin VB.PictureBox picMain 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   3
         Top             =   120
         Width           =   2880
         Begin VB.PictureBox picIcons 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   1200
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CheckBox chkIcon 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Load"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   4
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblIcon 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Icon Info"
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   0
            Left            =   960
            TabIndex        =   6
            Top             =   750
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCan_Click()
Call Unload(Me)
End Sub

Private Sub cmdLoad_Click()
Dim l As New Collection
Dim i As Integer
For i = 0 To chkIcon.Count - 1
 If chkIcon(i).value = 1 Then l.Add i
Next i
If l.Count <> 0 Then Call frmMain.LoadMultiIcons(l)
Call Unload(Me)
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To Temp_Icon.ifdCount - 1
 If i <> 0 Then
  Call Load(picIcons(i))
  picIcons(i).Visible = True
  picIcons(i).Top = picIcons(i - 1).Top + picIcons(i - 1).Height + chkIcon(i - 1).Height + lblIcon(i - 1).Height + 4
  Set picIcons(i).Picture = Temp_Icon.ifdIcon(i)
  Call Load(chkIcon(i))
  chkIcon(i).Visible = True
  chkIcon(i).Top = picIcons(i).Top + picIcons(i).Height
  Call Load(lblIcon(i))
  lblIcon(i).Visible = True
  lblIcon(i).Top = chkIcon(i).Top + chkIcon(i).Height
  picMain.Height = lblIcon(i).Top + lblIcon(i).Height + 10
 Else
  Set picIcons(i).Picture = Temp_Icon.ifdIcon(i)
 End If

picIcons(i).Left = (picBack.Width / 2) - (picIcons(i).Width / 2) - 12

 chkIcon(i).value = 1
 lblIcon(i).Caption = "Width: " & Temp_Icon.ifdIconData(i).idWidth & vbCrLf & _
                      "Height: " & Temp_Icon.ifdIconData(i).idHeight & vbCrLf & _
                      "Colors: " & Temp_Icon.ifdIconData(i).idColorCount
 
 If lblIcon(i).Top + lblIcon(i).Height > picBack.Height Then vs.Max = picMain.Height - picBack.Height: vs.value = 0
Next i
End Sub

Private Sub vs_Change()
picMain.Top = -vs.value
End Sub
