Attribute VB_Name = "modGeneral"
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type

Public Function FileExist(ByVal FileName As String) As Boolean
On Error GoTo 1
Dim a As Long
a = FileLen(FileName)
FileExist = True
Exit Function
1
FileExist = False
End Function

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
Dim strBuffer As String
 strBuffer = String(750, Chr(0))
 Key$ = LCase$(Key$)
 GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Function GetRGB(ByVal CVal As Long) As COLORRGB
'returns rgb values
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function

Public Function FindObject(ByRef frm As Form, ByVal ControlName As String) As Object
'returns an object with the same name as ControlName
Dim i As Integer
Dim s As String, d As Integer
d = -1
If InStr(ControlName, ":") Then 'if object is apart of an control array
 s = Left(ControlName, InStr(ControlName, ":") - 1)
 d = CInt(Mid(ControlName, InStr(ControlName, ":") + 1))
Else
 s = ControlName
End If
For i = 0 To frm.Controls.Count - 1 'loop through the controls of the form and find the object
 If frm.Controls(i).Name = s Then
  If d <> -1 Then 'object is an array
   If frm.Controls(i).Index = d Then
    Set FindObject = frm.Controls(i)
    Exit Function
   End If
  Else 'object is not array
   Set FindObject = frm.Controls(i)
   Exit Function
  End If
 End If
Next i
End Function

Public Function FindWindow(ByVal FormName As String) As Form
'returns a form with the same name as FormName
Dim i As Integer
For i = 0 To Forms.Count - 1 'loop through loaded forms
 If Forms(i).Name = FormName Then
   Set FindWindow = Forms(i)
   Exit Function
 End If
Next i
End Function

Public Function LoadLanguage(ByVal FileName As String)
'dynamically loads a language pack
On Error Resume Next 'just in case there is a misspelling don't let it crash the program
Dim arr() As String, i As Integer, sFont As String, iFont As Integer
Dim s As String, l As Long, obj As Object, frm As Form
l = FreeFile()
Open FileName For Input As #l
 s = Input(LOF(l), #l)
Close #l
arr() = Split(s, vbCrLf)
For i = 0 To UBound(arr())
 If Left(arr(i), 2) <> "//" Then 'make sure it's not a comment
  l = InStr(arr(i), "=")
  If l <> 0 Then 'run through the file and seperate the objects name and value
   Select Case LCase(Left(arr(i), l - 1))
    Case "form" 'find the form that holds the controls
     Set frm = FindWindow(Mid(arr(i), l + 1))
    Case "fontname" 'set the fontname
     sFont = Mid(arr(i), l + 1)
    Case "fontsize" 'set the fontsize
     iFont = CInt(Mid(arr(i), l + 1))
    Case Else
     Set obj = FindObject(frm, Left(arr(i), l - 1)) 'find the object
     obj.Caption = Mid(arr(i), l + 1) 'set it's caption
     obj.Text = Mid(arr(i), l + 1) 'set it's text
     obj.FontName = sFont 'these will generate errors on some controls
     obj.FontSize = iFont 'but we'll resume next
   End Select
  End If
 End If
Next i
End Function

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
 Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
 DoEvents
End Sub

