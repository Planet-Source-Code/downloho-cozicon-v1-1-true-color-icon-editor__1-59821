Attribute VB_Name = "modFilter"
Public Function Eval(expr As String) As Double
'originally written by Aldo Vargas and modified to fit CozIcon's needs
    Dim value As Variant, operand As String
    Dim pos As Integer
    
    pos = 1

If IsNumeric(expr) = True Then Eval = expr: Exit Function
    Do Until pos > Len(expr)

    Select Case Mid(expr, pos, 1)
        Case " "
        pos = pos + 1
        Case "+", "-", "*", "/", "\", "^"
        operand = Mid(expr, pos, 1)
        pos = pos + 1
        Case ">", "<", "=":

        operand = Mid(expr, pos, 1)

    pos = pos + 1
    Case Else

    Select Case operand
        Case "": value = Token(expr, pos)
        Case "+": Eval = Eval + value
        value = Token(expr, pos)
        Case "-": Eval = Eval + value
        value = -Token(expr, pos)
        Case "*": value = value * Token(expr, pos)
        Case "/": value = value / Token(expr, pos)
        Case "\": value = value \ Token(expr, pos)
        Case "^": value = value ^ Token(expr, pos)
    End Select

End Select

Loop


Eval = Eval + value
End Function


Private Function Token(expr, pos)

    Dim char As String, value As String, fn As String
    Dim es As Integer, pl As Integer, arr() As String
    Const QUOTE As String = """"

    Do Until pos > Len(expr)
        char = Mid(expr, pos, 1)


        Select Case char
         Case "+", "-", "/", "\", "*", "^", " ": Exit Do
         Case "("
          pl = 1
          pos = pos + 1
          es = pos

        Do Until pl = 0 Or pos > Len(expr)
         char = Mid(expr, pos, 1)

         Select Case char
          Case "(": pl = pl + 1
          Case ")": pl = pl - 1
         End Select

         pos = pos + 1
        Loop

        value = Mid(expr, es, pos - es - 1)
        fn = LCase(Token)

        Select Case fn
         Case "sin": Token = Sin(Eval(value))
         Case "cos": Token = Cos(Eval(value))
         Case "tan": Token = Tan(Eval(value))
         Case "exp": Token = Exp(Eval(value))
         Case "log": Token = Log(Eval(value))
         Case "atn": Token = Atn(Eval(value))
         Case "abs": Token = Abs(Eval(value))
         Case "sgn": Token = Sgn(Eval(value))
         Case "sqr": Token = Sqr(Eval(value))
         Case "pi": Token = 3.141592654
         Case "rgb"
          arr = Split(value, ">")
          Token = RGB(Eval(Trim(arr(0))), Eval(Trim(arr(1))), Eval(Trim(arr(2))))
         Case "rnd"
          Call Randomize
          Token = (Eval(value) * Rnd)
         Case "int": Token = Int(Eval(value))
         Case "red"
          arr = Split(value, ",")
          Token = GetRGB(frmMain.picIcon(frmMain.mCurrIcon).Point(Eval(Trim(arr(0))), Eval(Trim(arr(1))))).Red
         Case "green"
          arr = Split(value, ",")
          Token = GetRGB(frmMain.picIcon(frmMain.mCurrIcon).Point(Eval(Trim(arr(0))), Eval(Trim(arr(1))))).Green
         Case "blue"
          arr = Split(value, ",")
          Token = GetRGB(frmMain.picIcon(frmMain.mCurrIcon).Point(Eval(Trim(arr(0))), Eval(Trim(arr(1))))).Blue
         Case Else: Token = Eval(value)
        End Select

    Exit Do
    Case Else
    Token = Token & char
    pos = pos + 1
End Select

Loop
If IsNumeric(Token) Then Token = Val(Token)
End Function
