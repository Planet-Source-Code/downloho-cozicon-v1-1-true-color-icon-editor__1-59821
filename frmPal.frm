VERSION 5.00
Begin VB.Form frmPal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Icon Palette"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picClr 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   4080
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox picPal 
      AutoRedraw      =   -1  'True
      Height          =   3900
      Left            =   120
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   120
      Width           =   3900
   End
End
Attribute VB_Name = "frmPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Me.Hide
End Sub

Private Sub picPal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picClr.BackColor = picPal.Point(X, Y)
End Sub
