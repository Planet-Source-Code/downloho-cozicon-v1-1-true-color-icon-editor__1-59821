VERSION 5.00
Begin VB.UserControl CustomBrushes 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.VScrollBar vs 
      Height          =   855
      LargeChange     =   24
      Left            =   4560
      Max             =   0
      SmallChange     =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picDefault 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   360
      Picture         =   "CustomBrushes.ctx":0000
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   2
      Top             =   0
      Width           =   975
      Begin VB.PictureBox picPre 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -360
         MouseIcon       =   "CustomBrushes.ctx":04F2
         MousePointer    =   99  'Custom
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
   End
End
Attribute VB_Name = "CustomBrushes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum BT_BRUSHTYPE
 btPen = 0
 btBox = 1
 btNull = 99
End Enum

Const m_def_Path = ""
'My variables:
Dim Brushes() As String
Dim BrushPath() As String
'Property Variables:
Dim m_Path As String
Dim m_BrushType As BT_BRUSHTYPE
Dim m_CTop As Integer
Dim m_CLeft As Integer
'Event Declarations:
Event Click(ByVal Index As Integer)

Private Function DrawPreview(ByVal Data As String, ByVal FileName As String)
Dim p As PictureBox, l As Integer

 l = picPre.Count
  Call Load(picPre(l))
  With picPre(l)
   .Visible = True
   .Left = m_CLeft
   .Top = m_CTop
   
   If m_CLeft + 48 > UserControl.ScaleWidth - vs.Width Then
    m_CTop = m_CTop + 24
    m_CLeft = 4
   Else
    m_CLeft = m_CLeft + 24
   End If

   .ToolTipText = Left(FileName, Len(FileName) - 4)
  End With
  If m_CTop + 24 >= picBack.ScaleHeight Then picBack.Height = m_CTop + 24
  Set p = picPre(l)

Dim x As Integer, y As Integer
For l = 1 To Len(Data)
 Select Case Mid(Data, l, 1)
  Case "."
   p.Line (x, y)-(x + 4, y + 4), 0, BF
  Case ","
   p.Line (x, y)-(x + 4, y + 4), vbWhite, BF
  Case "'"
   p.Line (x, y)-(x + 4, y + 4), RGB(192, 192, 192), BF
  End Select
 x = x + 4
 If x = 20 Then x = 0: y = y + 4
Next l

p.Refresh
End Function

Public Function GetBrushData(ByVal Index As Integer) As String
 GetBrushData = Brushes(Index)
End Function

Private Sub LoadBrush(ByVal Path As String, ByVal Index As Integer)
Dim s As String, l As Long
l = FreeFile()
Open m_Path & Path For Input As #l
 s = Input(LOF(l), #l)
Close #l
Dim b As BT_BRUSHTYPE
b = btNull
If LCase(Left(s, 4)) = "draw" Then b = btPen
If LCase(Left(s, 3)) = "box" Then b = btBox

If m_BrushType <> b Then Exit Sub

If InStr(s, "<pre>") <> 0 And InStr(s, "</pre>") <> 0 Then
 p = Mid(s, InStr(s, "<pre>") + 5, InStr(InStr(s, "<pre>") + 1, s, "</pre>") - InStr(s, "<pre>") - 5)
 s = Left(s, InStr(s, "<pre>") - 1)
End If

If Len(p) = 25 Then
 'draw preview
 'Debug.Print p, Path
 Call DrawPreview(p, Path)
Else
 If p <> "" Then MsgBox "The preview for this Brush is incomplete!" & vbCrLf & Len(p) & " - " & p & " - " & Path, vbInformation, "Preview Error"
 'show default
 l = picPre.Count
  Call Load(picPre(l))
  With picPre(l)
   .Visible = True
   .Left = 4 + picPre(l - 1).Left + picPre(l - 1).Width
   .ToolTipText = Path
   Set .Picture = picDefault.Picture
  End With
End If
'MsgBox s
ReDim Preserve Brushes(Index)
Brushes(Index) = s
End Sub

Private Sub SortArray(ByRef arr() As String)
Dim bSrt As Boolean, temp As String, i As Integer

Do
bSrt = True

 For i = LBound(arr) To UBound(arr) - 1
    If IsNumeric(arr(LBound(arr))) = False Then
     If Asc(LCase(Left(arr(i), 1))) > Asc(LCase(Left(arr(i + 1), 1))) Then
      bSrt = False
      temp = arr(i)
      arr(i) = arr(i + 1)
      arr(i + 1) = temp
     End If
    Else
     If CLng(arr(i)) < CLng(arr(i + 1)) Then
      bSrt = False
      temp = arr(i)
      arr(i) = arr(i + 1)
      arr(i + 1) = temp
     End If
    End If
 Next i
Loop Until bSrt = True
End Sub

Public Property Get BrushType() As BT_BRUSHTYPE
    BrushType = m_BrushType
End Property

Public Property Let BrushType(b As BT_BRUSHTYPE)
    m_BrushType = b
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Path(ByVal New_Path As String)
    m_Path = New_Path
    PropertyChanged "Path"
   Dim sT As String, lA As Integer, arr() As String
    sT = Dir(m_Path, vbDirectory)
    If sT = "" Then Exit Property
     Do
      If sT <> "." And sT <> ".." Then
       If (GetAttr(m_Path & sT) And vbDirectory) <> vbDirectory Then
        ReDim Preserve arr(lA)
        arr(lA) = sT
        lA = lA + 1
       End If
      End If
      sT = Dir$
     Loop Until sT = ""

Call SortArray(arr())
Dim lB As Integer
Dim s As String, l As Long

For lA = 0 To UBound(arr())

l = FreeFile()
Open m_Path & arr(lA) For Input As #l
 s = Input(LOF(l), #l)
Close #l
Dim b As BT_BRUSHTYPE
b = btNull
If LCase(Left(s, 4)) = "draw" Then b = btPen
If LCase(Left(s, 3)) = "box" Then b = btBox

If m_BrushType = b Then
 ReDim Preserve BrushPath(lB)
 BrushPath(lB) = arr(lA)
 Call LoadBrush(BrushPath(lB), lB)
 lB = lB + 1
End If
Next lA
End Property

Private Sub picBack_Resize()
If picBack.Height > UserControl.ScaleHeight Then
 vs.Max = picBack.Height - UserControl.ScaleHeight
Else
 vsmax = 0
End If
 vs.value = 0
End Sub

Private Sub picPre_Click(Index As Integer)
RaiseEvent Click(Index - 1)
End Sub

Private Sub UserControl_Initialize()
    m_CLeft = 4
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BrushType = btPen
    m_Path = m_def_Path
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Path = PropBag.ReadProperty("Path", m_def_Path)
    m_BrushType = PropBag.ReadProperty("Path", 0)
End Sub

Private Sub UserControl_Resize()
vs.Height = UserControl.ScaleHeight
vs.Left = UserControl.ScaleWidth - vs.Width
picBack.Width = UserControl.ScaleWidth
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Path", m_Path, m_def_Path)
    Call PropBag.WriteProperty("Brushtype", m_BrushType, 0)
End Sub

Private Sub vs_Change()
picBack.Top = -vs.value
End Sub
