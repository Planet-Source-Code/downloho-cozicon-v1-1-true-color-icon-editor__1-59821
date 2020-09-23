VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFilterMaker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CozIcon Filter Maker"
   ClientHeight    =   3105
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8550
   Icon            =   "frmFilterMaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Cozicon Custom Filter (*.ccf)|*.ccf"
      Flags           =   7
   End
   Begin VB.CheckBox chkOffAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   2280
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.TextBox txtO 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Text            =   "0"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Code"
      Height          =   3015
      Left            =   2040
      TabIndex        =   15
      Top             =   0
      Width           =   6495
      Begin FilterMaker.winConnect wc 
         Left            =   5520
         Top             =   2400
         _ExtentX        =   1535
         _ExtentY        =   661
      End
      Begin VB.PictureBox picUndo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   5520
         Picture         =   "frmFilterMaker.frx":0D4A
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   5520
         MouseIcon       =   "frmFilterMaker.frx":1A16
         MousePointer    =   99  'Custom
         Picture         =   "frmFilterMaker.frx":26E0
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "untitled"
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picTrans 
         BackColor       =   &H00808000&
         Height          =   180
         Left            =   6240
         MouseIcon       =   "frmFilterMaker.frx":33AC
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Transparent Color"
         Top             =   960
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.CommandButton cmdPre 
         Caption         =   "Preview"
         Height          =   375
         Left            =   5160
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCode 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtW 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Text            =   "0"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   7
      Left            =   720
      TabIndex        =   7
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Offset:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Will add/subtract this number from final calculation"
      Top             =   2280
      Width           =   465
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Scale:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Final Calculation will be divided by this number"
      Top             =   1920
      Width           =   450
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
   End
   Begin VB.Menu mnuSend 
      Caption         =   "Send to CozIcon"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmFilterMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGen_Click()
Dim s(8, 1) As String, d As String
Dim i As Integer
'if negative subtract (value * 2) if positive add (value * 2)
If txt(0).Text <> "0" Then
 s(0, 1) = "+"
 Select Case Left(txt(0).Text, 1)
  Case "-"
   s(0, 1) = "-"
   If Mid(txt(0).Text, 2) = "1" Then
    s(0, 0) = "red(x - 1, y - 1)"
   Else
    s(0, 0) = "(red(x - 1, y - 1) * " & Mid(txt(0).Text, 2) & ")"
   End If
  Case "1"
   s(0, 0) = "red(x - 1, y - 1)"
  Case Else
   s(0, 0) = "(red(x - 1, y - 1) * " & txt(0).Text & ")"
 End Select
End If

If txt(1).Text <> "0" Then
 s(1, 1) = "+"
 Select Case Left(txt(1).Text, 1)
  Case "-"
   s(1, 1) = "-"
   If Mid(txt(1).Text, 2) = "1" Then
    s(1, 0) = "red(x, y - 1)"
   Else
    s(1, 0) = "(red(x, y - 1) * " & Mid(txt(1).Text, 2) & ")"
   End If
  Case "1"
   s(1, 0) = "red(x, y - 1)"
  Case Else
   s(1, 0) = "(red(x, y - 1) * " & txt(1).Text & ")"
 End Select
End If

If txt(2).Text <> "0" Then
 s(2, 1) = "+"
 Select Case Left(txt(2).Text, 1)
  Case "-"
   s(2, 1) = "-"
   If Mid(txt(2).Text, 2) = "1" Then
    s(2, 0) = "red(x + 1, y - 1)"
   Else
    s(2, 0) = "(red(x + 1, y - 1) * " & Mid(txt(2).Text, 2) & ")"
   End If
  Case "1"
   s(2, 0) = "red(x + 1, y - 1)"
  Case Else
   s(2, 0) = "(red(x + 1, y - 1) * " & txt(2).Text & ")"
 End Select
End If

If txt(3).Text <> "0" Then
 s(3, 1) = "+"
 Select Case Left(txt(3).Text, 1)
  Case "-"
   s(3, 1) = "-"
   If Mid(txt(3).Text, 2) = "1" Then
    s(3, 0) = "red(x - 1, y)"
   Else
    s(3, 0) = "(red(x - 1, y) * " & Mid(txt(3).Text, 2) & ")"
   End If
  Case "1"
   s(3, 0) = "red(x - 1, y)"
  Case Else
   s(3, 0) = "(red(x - 1, y) * " & txt(3).Text & ")"
 End Select
End If

If txt(4).Text <> "0" Then
 s(4, 1) = "+"
 Select Case Left(txt(4).Text, 1)
  Case "-"
   s(4, 1) = "-"
   If Mid(txt(4).Text, 2) = "1" Then
    s(4, 0) = "red(x, y)"
   Else
    s(4, 0) = "(red(x, y) * " & Mid(txt(4).Text, 2) & ")"
   End If
  Case "1"
   s(4, 0) = "red(x, y)"
  Case Else
   s(4, 0) = "(red(x, y) * " & txt(4).Text & ")"
 End Select
End If

If txt(5).Text <> "0" Then
 s(5, 1) = "+"
 Select Case Left(txt(5).Text, 1)
  Case "-"
   s(5, 1) = "-"
   If Mid(txt(5).Text, 2) = "1" Then
    s(5, 0) = "red(x + 1, y)"
   Else
    s(5, 0) = "(red(x + 1, y) * " & Mid(txt(5).Text, 2) & ")"
   End If
  Case "1"
   s(5, 0) = "red(x + 1, y)"
  Case Else
   s(5, 0) = "(red(x + 1, y) * " & txt(5).Text & ")"
 End Select
End If

If txt(6).Text <> "0" Then
 s(6, 1) = "+"
 Select Case Left(txt(6).Text, 1)
  Case "-"
   s(6, 1) = "-"
   If Mid(txt(6).Text, 2) = "1" Then
    s(6, 0) = "red(x - 1, y + 1)"
   Else
    s(6, 0) = "(red(x - 1, y + 1) * " & Mid(txt(6).Text, 2) & ")"
   End If
  Case "1"
   s(6, 0) = "red(x - 1, y + 1)"
  Case Else
   s(6, 0) = "(red(x - 1, y + 1) * " & txt(6).Text & ")"
 End Select
End If

If txt(7).Text <> "0" Then
 s(7, 1) = "+"
 Select Case Left(txt(7).Text, 1)
  Case "-"
   s(7, 1) = "-"
   If Mid(txt(7).Text, 2) = "1" Then
    s(7, 0) = "red(x, y + 1)"
   Else
    s(7, 0) = "(red(x, y + 1) * " & Mid(txt(7).Text, 2) & ")"
   End If
  Case "1"
   s(7, 0) = "red(x, y + 1)"
  Case Else
   s(7, 0) = "(red(x, y + 1) * " & txt(7).Text & ")"
 End Select
End If

If txt(8).Text <> "0" Then
 s(8, 1) = "+"
 Select Case Left(txt(8).Text, 1)
  Case "-"
   s(8, 1) = "-"
   If Mid(txt(8).Text, 2) = "1" Then
    s(8, 0) = "red(x + 1, y + 1)"
   Else
    s(8, 0) = "(red(x + 1, y + 1) * " & Mid(txt(8).Text, 2) & ")"
   End If
  Case "1"
   s(8, 0) = "red(x + 1, y + 1)"
  Case Else
   s(8, 0) = "(red(x + 1, y + 1) * " & txt(8).Text & ")"
 End Select
End If

For i = 0 To 8
 If s(i, 0) <> "" Then
  d = d & " " & s(i, 1) & " " & s(i, 0)
 End If
 'If i = 2 And d <> "" Then d = Left(d, Len(d) - 3) & " * "
Next i

If d <> "" Then d = Mid(d, 4)
d = "(" & d & ")"

d = "red = " & d & IIf(txtW.Text <> "0" And txtW.Text <> "1" And IsNumeric(txtW.Text) = True, " / " & txtW.Text, "")

If txtO.Text <> 0 And IsNumeric(txtO.Text) = True Then
 If chkOffAdd.value = 1 Then
  d = d & " + " & txtO.Text
 Else
  d = d & " - " & txtO.Text
 End If
End If
txtCode.Text = "filter" & vbCrLf & "#*" & vbCrLf & d & vbCrLf & Replace(d, "red", "green") & vbCrLf & Replace(d, "red", "blue")
End Sub

Private Sub Label1_Click()

End Sub

Private Function KillFile(ByVal file As String) As Boolean
On Error GoTo 1
Call Kill(file)
1
End Function

Private Sub cmdPre_Click()
If txtCode.Text = "" Then Exit Sub
Dim l As Long, x As Integer, y As Integer
Dim s As String, arr() As String, a As Integer

 s = txtCode.Text

arr() = Split(s, vbCrLf)
If arr(0) <> "filter" Then Exit Sub
For l = 1 To UBound(arr())
 If arr(l) = "#" & picIcon.ScaleWidth Or arr(l) = "#*" Then a = l + 1: Exit For
Next l

Dim c As Integer, d As Long
Dim sL As String, sR As String
Dim r As Double, g As Double, b As Double

Dim pX As Integer, pY As Integer
Dim sX As Integer, sY As Integer
Dim stX As Integer, stY As Integer
Dim UserInput As String

For l = a To UBound(arr())
 If Len(arr(l)) < 10 Then a = l: Exit For
 Select Case LCase(Left(arr(l), 10))
  Case "user_input"
   UserInput = Trim(InputBox(Mid(arr(l), 12), "User Input"))
   If UserInput = "" Then Exit Sub
  Case "pixelskipx"
   pX = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelskipy"
   pY = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelstrtx"
   sX = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelstrty"
   sY = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelstopx"
   stX = CInt(Trim(Mid(arr(l), 12)))
  Case "pixelstopy"
   stY = CInt(Trim(Mid(arr(l), 12)))
  Case Else
   a = l
   Exit For
 End Select
Next l
stY = IIf(stY = 0, picIcon.ScaleHeight - 1, stY)
stX = IIf(stX = 0, picIcon.ScaleHeight - 1, stX)

frmFilter.Show
frmFilter.Refresh
frmFilter.ProgressBar1.Max = stY

For y = sY To stY
 For x = sX To stX
  For l = a To UBound(arr())
   If Left(arr(l), 1) = "#" Then Exit For
    c = InStr(arr(l), "=")
    'Debug.Print picIcon(mCurrIcon).Point(X, Y), picTrans.BackColor, X, Y
    If c <> 0 And picIcon.Point(x, y) <> picTrans.BackColor Then
     sL = Trim(Left(arr(l), c - 1))
     sR = Trim(Mid(arr(l), c + 1))
     If UserInput <> "" And IsNumeric(UserInput) = True Then sR = Trim(Replace(sR, "ui", CDbl(UserInput), , , vbTextCompare))
     sR = Trim(Replace(sR, "x", x, , , vbTextCompare))
     sR = Trim(Replace(sR, "y", y, , , vbTextCompare))
     sR = Trim(Replace(sR, "w", picIcon.ScaleWidth, , , vbTextCompare))
     sR = Trim(Replace(sR, "h", picIcon.ScaleHeight, , , vbTextCompare))
     'Debug.Print sR, Eval(sR)
      Select Case LCase(sL)
       Case "red"
        r = Abs(Eval(sR))
       Case "green"
        g = Abs(Eval(sR))
       Case "blue"
        b = Abs(Eval(sR))
      End Select
      If r > 255 Then r = 255
      If g > 255 Then g = 255
      If b > 255 Then b = 255
      If RGB(r, g, b) = picTrans.BackColor Then r = r + 1: g = g + 1: b = b + 1
            picIcon.PSet (x, y), RGB(r, g, b)
    End If
  Next l
   x = x + pX
 Next x
 frmFilter.ProgressBar1.value = y
 y = y + pY
Next y
frmFilter.Hide

Set picIcon.Picture = picIcon.Image
End Sub

Private Sub Form_Load()
If Command <> "" Then
 Call wc.SetCommand(Command)
 mnuSend.Visible = True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub mnuNew_Click()
Dim i As Integer
For i = 0 To 8
 txt(i).Text = 0
Next i

txtW.Text = 0
txtO.Text = 0
chkOffAdd.value = 1
txtCode.Text = "filter" & vbCrLf & "#*" & vbCrLf
End Sub

Private Sub mnuSave_Click()
On Error GoTo 1
cd.Filter = "Cozicon Custom Filter (*.ccf)|*.ccf|Filter Maker File (*.fmf)|*.fmf"
cd.ShowSave

If cd.FileName = "" Then Exit Sub
Dim s As String
Select Case LCase(Right(cd.FileName, 3))
 Case "fmf"
  Dim i As Integer
  
  For i = 0 To 8
   s = s & txt(0).Text & "|"
  Next i
   s = s & txtW.Text & "|" & txtO.Text & "|" & chkOffAdd.value & "|" & txtCode.Text
  
   Call KillFile(cd.FileName)
   Open cd.FileName For Binary Access Write As #1
    Put #1, , s$
   Close #1
 Case Else
  s = txtCode.Text
   Call KillFile(cd.FileName)
   Open cd.FileName For Binary Access Write As #1
    Put #1, , s$
   Close #1
End Select
1
End Sub

Private Sub mnuSend_Click()
Set picIcon.Picture = picIcon.Image
Call SavePicture(picIcon.Picture, picUndo.Tag)
Call wc.Send("!")
End
End Sub

Private Sub picIcon_Click()
If picUndo.Tag <> "" Then Exit Sub
On Error GoTo 1
cd.Filter = "Icons (*.ico)|*.ico"
cd.ShowOpen

If cd.FileName = "" Then Exit Sub

Set picUndo.Picture = LoadPicture(cd.FileName)
Set picIcon.Picture = LoadPicture(cd.FileName)
1
End Sub

Private Sub txt_GotFocus(Index As Integer)
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index).Text)
End Sub

Private Sub TxtO_GotFocus()
txtO.SelStart = 0
txtO.SelLength = Len(txtO.Text)
End Sub

Private Sub txtW_GotFocus()
txtW.SelStart = 0
txtW.SelLength = Len(txtW.Text)
End Sub

Private Sub wc_Got(ByVal Msg As String)
Select Case Left(Msg, 1)
 Case "$"
  picUndo.Tag = Mid(Msg, 2)
  Set picIcon.Picture = LoadPicture(Mid(Msg, 2))
  Set picUndo.Picture = picIcon.Picture
  Me.Show
End Select
End Sub

