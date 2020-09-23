VERSION 5.00
Begin VB.UserControl ColorPicker 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Event Declarations:
Event Pick(ByVal Button As Integer, ByVal Clr As Long)
Dim mClrs() As Long
'Default Property Values:
Const m_def_SelWidth = 16
Const m_def_SelHeight = 16
'Property Variables:
Dim m_SelWidth As Integer
Dim m_SelHeight As Integer

Public Sub AddColor(ByVal Clr As Long)
Dim i As Integer
i = UBound(mClrs()) + 1
ReDim Preserve mClrs(i)

mClrs(i) = Clr
End Sub

Public Sub ClearColors()
ReDim mClrs(0)
End Sub

Private Sub DrawInvertedBox(ByVal X As Integer, ByVal Y As Integer, ByVal W As Integer, ByVal H As Integer, ByVal BackClr As Long)
Line (X, Y)-(X + W, Y + H), BackClr, BF

Line (X, Y)-(X + W, Y), 0 'top high
Line (X, Y + 1)-(X + W - 1, Y + 1), RGB(128, 128, 128)

Line (X, Y)-(X, Y + H), 0 'left high
Line (X + 1, Y + 1)-(X + 1, Y + H - 1), RGB(128, 128, 128)

Line (X + W, Y)-(X + W, Y + H + 1), vbWhite 'right shadow
Line (X + W - 1, Y + 1)-(X + W - 1, Y + H), RGB(192, 192, 192)  'right shadow

Line (X, Y + H)-(X + W, Y + H), vbWhite 'bottom shadow
Line (X, Y + H - 1)-(X + W - 1, Y + H - 1), RGB(192, 192, 192) 'bottom shadow

Refresh
End Sub

Public Sub ShowColors()
Dim i As Integer
Dim X As Integer
Dim Y As Integer

For i = 1 To UBound(mClrs())
Debug.Print i, mClrs(i)
 Call DrawInvertedBox(X, Y, m_SelWidth, m_SelHeight, mClrs(i))
 X = X + m_SelWidth
 If X + m_SelWidth > ScaleWidth Then X = 0: Y = Y + m_SelHeight
 Debug.Print X, Y
Next i
Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Initialize()
ReDim mClrs(0)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 And Button <> 2 Then Exit Sub
RaiseEvent Pick(Button, Point(Int(X / m_SelWidth) * m_SelWidth + 4, Int(Y / m_SelHeight) * m_SelHeight + 4))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 And Button <> 2 Then Exit Sub
RaiseEvent Pick(Button, Point(Int(X / m_SelWidth) * m_SelWidth + 4, Int(Y / m_SelHeight) * m_SelHeight + 4))
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_SelWidth = PropBag.ReadProperty("SelWidth", m_def_SelWidth)
    m_SelHeight = PropBag.ReadProperty("SelHeight", m_def_SelHeight)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("SelWidth", m_SelWidth, m_def_SelWidth)
    Call PropBag.WriteProperty("SelHeight", m_SelHeight, m_def_SelHeight)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,16
Public Property Get SelWidth() As Integer
    SelWidth = m_SelWidth
End Property

Public Property Let SelWidth(ByVal New_SelWidth As Integer)
    m_SelWidth = New_SelWidth
    PropertyChanged "SelWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,16
Public Property Get SelHeight() As Integer
    SelHeight = m_SelHeight
End Property

Public Property Let SelHeight(ByVal New_SelHeight As Integer)
    m_SelHeight = New_SelHeight
    PropertyChanged "SelHeight"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SelWidth = m_def_SelWidth
    m_SelHeight = m_def_SelHeight
End Sub

