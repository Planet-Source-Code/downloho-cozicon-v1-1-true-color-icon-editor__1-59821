VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Brush Preview Maker"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picClr 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   2
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Tag             =   "'"
      Top             =   3600
      Width           =   375
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Tag             =   ","
      Top             =   3600
      Width           =   375
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Tag             =   "."
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox txtPre 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate Preview"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.PictureBox picPre 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   120
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mClr As Integer

Private Sub cmdGen_Click()
Dim s As String

For Y = 0 To 4
 For X = 0 To 4
  c = picPre.Point(X * 40 + 1, Y * 40 + 1)
  Select Case c
   Case 0
    s = s & "."
   Case vbWhite
    s = s & ","
   Case RGB(192, 192, 192)
    s = s & "'"
   Case Else
    s = s & " "
  End Select
    
 Next X
Next Y

txtPre.Text = "<pre>" & s & "</pre>"
End Sub

Private Sub picClr_Click(Index As Integer)
mClr = Index
End Sub

Private Sub picPre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picPre.Point(X, Y) <> 0 And picPre.Point(X, Y) <> vbWhite And picPre.Point(X, Y) <> RGB(192, 192, 192) Then
 Select Case mClr
  Case 0
   picPre.Line (Int(X / 40) * 40, Int(Y / 40) * 40)-(Int(X / 40) * 40 + 39, Int(Y / 40) * 40 + 39), 0, BF
  Case 1
   picPre.Line (Int(X / 40) * 40, Int(Y / 40) * 40)-(Int(X / 40) * 40 + 39, Int(Y / 40) * 40 + 39), vbWhite, BF
  Case 2
   picPre.Line (Int(X / 40) * 40, Int(Y / 40) * 40)-(Int(X / 40) * 40 + 39, Int(Y / 40) * 40 + 39), RGB(192, 192, 192), BF
 End Select
Else
 picPre.Line (Int(X / 40) * 40, Int(Y / 40) * 40)-(Int(X / 40) * 40 + 39, Int(Y / 40) * 40 + 39), RGB(0, 128, 128), BF
End If
End Sub

Private Sub picPre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = Int(X / 40) & ", " & Int(Y / 40)
End Sub
