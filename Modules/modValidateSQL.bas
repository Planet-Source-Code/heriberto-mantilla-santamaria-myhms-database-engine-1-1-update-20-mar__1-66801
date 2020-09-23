Attribute VB_Name = "modValidateSQL"
Option Explicit
'By DreamVB
'This is my Recursive Decent Parsing project. I made this project to maybe help others
'That may need to use something like this in there projects.
'I made this after reading though one of my C++ books and decided to convert my C++ project over to
'   VB

'Token Types

Private Enum token_type
   LERROR = -1
   NONE = 0
   DELIMITER = 1
   DIGIT = 2
   LSTRING = 3
   VARIABLE = 4
   IDENTIFIER = 6
   HEXDIGIT = 5
   FINISHED = 7
End Enum

'Relational
Private Const GE = 1 ' Greator than or equal to
Private Const NE = 2 ' not equal to
Private Const LE = 3 ' Less than or equal to

'Bitwise
Private Const cAND = 4
Private Const cOR = 5
'Bitshift

Private Const shr = 6
Private Const shl = 7
Private Const cXor = 8
Private Const cIMP = 9
Private Const cEqv = 11
Private Const cINC = 12

Private Const Str_Ops = "AND,OR,XOR,MOD,DIV,SHL,SHR,IMP,EQV,NOT"
Private Const Str_Funcs = "ABS,ATN,COS,EXP,LOG,RND,ROUND,SGN,SIN,SQR,TAN,SUM,IIF"

'We use this to store variables

Private Type vars
   vName As String
   vValue As Variant
End Type

Private Token As String         'Current processing token
Private tok_type As token_type  'Used to idenfiy the tokens
Private Look_Pos As Long    'Current processing char pointer
Private ExprLine As String  'The Expression line to scan

Private lVars() As vars   '26 variables
Private lVarCount As Integer

Public isAbort As Boolean

Private Sub Abort(code As Integer, Optional aStr As String = "")

  Dim lMsg As String

   Select Case code
    Case 0: lMsg = "Undeclared variable '" & aStr & "'"
    Case 1: lMsg = "Division by zero"
    Case 2: lMsg = "Expected parenthesized closing bracket ')'"
    Case 3: lMsg = "Invalid Digit found '" & aStr & "'"
    Case 4: lMsg = "Unknown character found '" & aStr & "'"
    Case 5: lMsg = "The Variable '" & aStr & "' is an identifier and can't be used."
    Case 6: lMsg = "Expected expression."
    Case 7: lMsg = "Invalid Hexadecimal value found '0x" & UCase(aStr) & "'"
   End Select

   isAbort = True

   Look_Pos = Len(ExprLine) + 1

End Sub

Public Sub AddVar(Name As String, Optional lValue As Variant = 0)

   'Add a new variable along with the variables value.

   ReDim Preserve lVars(lVarCount)     'Resize variable stack
   lVars(lVarCount).vName = Name       'Add variable name
   lVars(lVarCount).vValue = lValue    'Add varaible data
   lVarCount = lVarCount + 1           'INC variable Counter

End Sub

Private Function atom()

  Dim Temp As String

   'Check for Digits ,Hexadecimal,Functions, Variables

   Select Case tok_type
    Case HEXDIGIT 'Hexadecimal
      Temp = Trim$(Right(Token, Len(Token) - 2))

      If Len(Temp) = 0 Then
         Abort 6
       ElseIf Not isHex(Temp) Then
         Abort 7, Temp
       Else
         atom = CDec("&H" & Temp)
         GetToken
      End If

    Case IDENTIFIER 'Inbuilt Functions
      atom = CallIntFunc(Token)
      GetToken

    Case DIGIT 'Digit const found
      If Not IsNumeric(Token) Then Abort 3, Token 'Check we have a real digit
      atom = Token 'Return the value
      GetToken 'Get next token

    Case LERROR 'Expression phase error
      Abort 0, Token 'Show error message

    Case VARIABLE 'Variable found
      If FindVarIdx(Token) = -1 Then Abort 0, Token
      atom = GetVarData(Token) 'Return variable value
      GetToken 'Get next token
   End Select

End Function

Private Function CallIntFunc(sFunction As String) As Double

  Dim Temp
  Dim UserFuncID As Integer
  Dim x As Integer
  Dim sFunction_Str As String
  Dim ArgList

   'ABS,ATN,COS,EXP,LOG,RND,ROUND,SGN,SIN,SQR,TAN,IFF

   On Error Resume Next

   Select Case UCase(sFunction)
    Case "ABS"
      GetToken
      Temp = Exp6
      CallIntFunc = Abs(Temp)
      PushBack

    Case "ATN"
      GetToken
      Temp = Exp6
      CallIntFunc = Atn(Temp)
      PushBack

    Case "COS"
      GetToken
      Temp = Exp6
      CallIntFunc = Cos(Temp)
      PushBack

    Case "EXP"
      GetToken
      Temp = Exp6
      CallIntFunc = Exp(Temp)
      PushBack

    Case "LOG"
      GetToken
      Temp = Exp6
      CallIntFunc = Log(Temp)
      PushBack

    Case "RND"
      GetToken
      Temp = Exp6
      CallIntFunc = Rnd(Temp)
      PushBack

    Case "ROUND"
      GetToken
      Temp = Exp6
      CallIntFunc = Round(Temp)
      PushBack

    Case "SGN"
      GetToken
      Temp = Exp6
      CallIntFunc = Sgn(Temp)
      PushBack

    Case "SIN"
      GetToken
      Temp = Exp6
      CallIntFunc = Sin(Temp)
      PushBack

    Case "SQR"
      GetToken
      Temp = Exp6
      CallIntFunc = Sqr(Temp)
      PushBack

    Case "TAN"
      GetToken
      Temp = Exp6
      CallIntFunc = Tan(Temp)
      PushBack

    Case "SUM"
      ArgList = GetArgs
      Temp = 0

      For x = 0 To UBound(ArgList)
         Temp = CDbl(Temp) + CDbl(ArgList(x))
      Next x

      GetToken
      CallIntFunc = Temp
      PushBack

    Case "IIF"
      ArgList = GetArgs
      CallIntFunc = IIf(ArgList(0), ArgList(1), ArgList(2))
      PushBack
   End Select

End Function

Public Function Eval(Expression As String)

On Error Resume Next
   isAbort = False
   ExprLine = Expression 'Store the expression to scan
   Look_Pos = 1    'Default state of char pos
   GetToken        'Kick start and Get the first token.

   If (tok_type = FINISHED) Or Len(Trim(Expression)) = 0 Then
      isAbort = True
      Exit Function
    Else
      Eval = Exp0()
   End If
On Error GoTo 0

End Function

Private Function Exp0()

  Dim Tmp_tokType As token_type
  Dim Tmp_Token As String
  Dim Var_Idx As Integer
  Dim Temp

   'Assignments

   If (tok_type = VARIABLE) Then
      'Store temp type and token
      'we first need to check if the variable name is not an identifier
      If isIdent(Token) Then Abort 5, Token
      Tmp_tokType = tok_type
      Tmp_Token = Token
      'Locate the variables index
      Var_Idx = FindVarIdx(Token)
      'If we have an invaild var index -1 we Must add a new variable

      If (Var_Idx = -1) Then
         'Add the new variable
         'AddVar Token
         'Now get the variable index agian
         'Var_Idx = FindVarIdx(Token)
         isAbort = True
         Exit Function
       Else
         Exp0 = Exp1
         Exit Function
      End If

      'Get the next token
      Call GetToken

      If (Token <> "=") Then
         PushBack 'Move expr pointer back
         Token = Tmp_Token       'Restore temp token
         tok_type = Tmp_tokType  'Restore temp token type
       Else
         'Carry on processing the expression
         Call GetToken
         'Set the variables value
         Temp = Exp1()
         SetVar Var_Idx, Temp
         Exp0 = Temp
      End If

   End If
   Exp0 = Exp1

End Function

Private Function Exp1()

  Dim op As String
  Dim Relops As String
  Dim Temp
  Dim rPos As Integer

   'Relational operators
   Relops = Chr$(GE) + Chr$(NE) + Chr$(LE) + "<" + ">" + "=" + "!" + Chr$(0)
   Exp1 = Exp2()

   op = Token 'Get operator
   rPos = InStr(1, Relops, op) 'Check for other ops in token <> =

   If rPos > 0 Then
      GetToken 'Get next token
      Temp = Exp2 'Store temp val

      Select Case op
       Case "<" 'less then
         Exp1 = CDbl(Exp1) < CDbl(Temp)

       Case ">" 'greator than
         Exp1 = CDbl(Exp1) > CDbl(Temp)

       Case Chr(NE)
         Exp1 = CDbl(Exp1) <> CDbl(Temp)

       Case Chr(LE)
         Exp1 = CDbl(Exp1) <= CDbl(Temp)

       Case Chr(GE)
         Exp1 = CDbl(Exp1) >= CDbl(Temp)

       Case "=" 'equal to
         Exp1 = Exp1 = Temp

       Case "!"
         Exp1 = Not CDbl(Temp)
      End Select

      'op = Token
   End If

End Function

Private Function Exp2()

  Dim op As String
  Dim Temp

   'Add or Subtact two terms
   Exp2 = Exp3()
   op = Token 'Get operator

   Do While (op = "+" Or op = "-")
      GetToken 'Get next token
      Temp = Exp3() 'Temp value
      'Peform the expresion for the operator

      Select Case op
       Case "-"
         Exp2 = CDbl(Exp2) - CDbl(Temp)

       Case "+"
         Exp2 = CDbl(Exp2) + CDbl(Temp)
      End Select

      op = Token
   Loop

End Function

Private Function Exp3()

  Dim op As String
  Dim Temp

   'Multiply or Divide two factors
   Exp3 = Exp4()
   op = Token 'Get operator

   Do While (op = "*" Or op = "/" Or op = "\" Or op = "%")
      GetToken 'Get next token
      Temp = Exp4() 'Temp value
      'Peform the expresion for the operator

      Select Case op
       Case "*"
         Exp3 = CDbl(Exp3) * CDbl(Temp)

       Case "/"
         If Temp = 0 Then Abort 1
         Exp3 = CDbl(Exp3) / CDbl(Temp)

       Case "\"
         If Temp = 0 Then Abort 1
         Exp3 = CDbl(Exp3) \ CDbl(Temp)

       Case "%"
         If Temp = 0 Then Abort 1
         Exp3 = CDbl(Exp3) Mod CDbl(Temp)
      End Select

      op = Token
   Loop

End Function

Private Function Exp4()

  Dim op As String
  Dim BitWOps As String
  Dim Temp
  Dim rPos As Integer

   'Bitwise operators ^ | & || &&
   BitWOps = Chr$(cAND) + Chr$(cOR) + Chr$(shl) + Chr$(shr) + Chr$(cXor) + Chr$(cIMP) + Chr$(cEqv) + "^" + _
      "|" + "&" + Chr(0)
   Exp4 = Exp5()

   op = Token 'Get operator
   rPos = InStr(1, BitWOps, op) 'Check for other ops in token <> =

   If rPos > 0 Then
      GetToken 'Get next token
      Temp = Exp5 'Store temp val

      Select Case op
       Case "^" 'Excompnent
         Exp4 = CDbl(Exp4) ^ CDbl(Temp)

       Case "&"
         Exp4 = CDbl(Exp4) & CDbl(Temp)

       Case Chr(cAND)
         Exp4 = CDbl(Exp4) And CDbl(Temp)

       Case Chr$(cOR)
         Exp4 = CDbl(Exp4) Or CDbl(Temp)

       Case Chr$(shl)
         'Bitshift Shift left
         Exp4 = CDbl(Exp4) * (2 ^ CDbl(Temp))

       Case Chr$(shr)
         'bitshift right
         Exp4 = CDbl(Exp4) \ (2 ^ CDbl(Temp))

       Case Chr$(cXor)
         'Xor
         Exp4 = CDbl(Exp4) Xor CDbl(Temp)

       Case Chr$(cIMP)
         'IMP
         Exp4 = CDbl(Exp4) Imp CDbl(Temp)

       Case Chr$(cEqv)
         Exp4 = CDbl(Exp4) Eqv CDbl(Temp)
      End Select

      'op = Token
   End If

End Function

Private Function Exp5()

  Dim op As String
  Dim Temp As Variant

   op = ""
   'Unary +,-

   If ((tok_type = DELIMITER) And (Token = "+" Or Token = "-")) Then
      op = Token
      GetToken
   End If

   Exp5 = Exp6()
   If (op = "-") Then Exp5 = -CDbl(Exp5)

End Function

Private Function Exp6()

   'Check for Parenthesized expression

   If Token = "(" Then
      GetToken 'Get next token
      Exp6 = Exp1()
      'Check that we have a closeing bracket
      If (Token <> ")") Then Abort 2
      GetToken 'Get next token
    Else
      Exp6 = atom()
   End If

End Function

Private Function FindVarIdx(Name As String) As Integer

  Dim x As Integer
  Dim idx As Integer

   'Locate a variables position in the variables array
   idx = -1 'Bad position

   For x = 0 To UBound(lVars)

      If LCase$(Name) = LCase$(lVars(x).vName) Then
         idx = x
         Exit For
      End If

   Next x
   FindVarIdx = idx

End Function

Private Function GetArgs()

  Dim Count As Integer
  Dim Value
  Dim Temp() As Variant

   GetToken
   If Token <> "(" Then Exit Function

   Do

      GetToken
      Value = Exp1
      ReDim Preserve Temp(0 To Count)
      Temp(Count) = Value
      Count = Count + 1

   Loop Until (Token = ")")

   GetArgs = Temp
   Erase Temp
   Count = 0
   Value = 0

End Function

Private Sub GetToken()

  Dim Temp As String
  Dim idx As Integer
  Dim dTmp

   Temp = ""
   'This is the main part of the pharser and is used to.
   'Identfiy all the tokens been scanned and return th correct token type

   'Clear current token info
   Token = ""
   tok_type = NONE

   If Look_Pos > Len(ExprLine) Then tok_type = FINISHED: Exit Sub
   'Above exsits the sub if we are passed expr len

   Do While (Look_Pos <= Len(ExprLine) And isWhite(Mid$(ExprLine, Look_Pos, 1)))
      'Skip over white spaces. and stay within the expr len
      Look_Pos = Look_Pos + 1 'INC
      If Look_Pos > Len(ExprLine) Then Exit Sub
   Loop

   'Some little test I was doing to do Increment/Decrement operators -- ++

   If ((Mid$(ExprLine, Look_Pos, 1) = "+") Or Mid$(ExprLine, Look_Pos, 1) = "-") Then
      If ((Mid$(ExprLine, Look_Pos + 1, 1) = "+") Or Mid$(ExprLine, Look_Pos + 1, 1) = "-") Then
         Temp = Mid(ExprLine, 1, Look_Pos - 1)

         If Mid$(ExprLine, Look_Pos + 1, 1) = "+" Then
            dTmp = GetVarData(Temp) + 1
          ElseIf Mid$(ExprLine, Look_Pos + 1, 1) = "-" Then
            dTmp = GetVarData(Temp) - 1
         End If

         SetVar FindVarIdx(Temp), dTmp
         Token = Temp
         Exit Sub
      End If

   End If
   ''

   If (Mid$(ExprLine, Look_Pos, 1) = "&") Or Mid$(ExprLine, Look_Pos, 1) = "|" Then
      'Bitwise code, I still got some work to do on this yet but it does the ones
      ' that are listed below fine

      Select Case Mid$(ExprLine, Look_Pos, 1)
       Case "&"

         If Mid$(ExprLine, Look_Pos + 1, 1) = "&" Then
            Look_Pos = Look_Pos + 2
            Token = Chr(cAND)
            Exit Sub
          Else
            Look_Pos = Look_Pos + 1
            Token = "&"
            Exit Sub
         End If

       Case "|"

         If Mid$(ExprLine, Look_Pos + 1, 1) = "|" Then
            Look_Pos = Look_Pos + 2
            Token = Chr$(cOR)
            Exit Sub
          Else
            Look_Pos = Look_Pos + 1
            Token = "|"
            Exit Sub
         End If

         tok_type = DELIMITER
      End Select

   End If

   If (Mid$(ExprLine, Look_Pos, 1) = "<") Or (Mid$(ExprLine, Look_Pos, 1) = ">") Then
      'Check for Relational operators < > <= >= <>
      'check for not equal to get first op <

      Select Case Mid$(ExprLine, Look_Pos, 1)
       Case "<"

         If Mid$(ExprLine, Look_Pos + 1, 1) = ">" Then
            'Not Equal to
            Look_Pos = Look_Pos + 2
            Token = Chr$(NE)
            Exit Sub
          ElseIf Mid$(ExprLine, Look_Pos + 1, 1) = "=" Then
            'Less then of equal to
            Look_Pos = Look_Pos + 2
            Token = Chr$(LE)
            Exit Sub
          ElseIf Mid$(ExprLine, Look_Pos + 1, 1) = "<" Then
            'Bitshift left
            Look_Pos = Look_Pos + 2
            Token = Chr$(shl)
            Exit Sub
          Else
            'Less then
            Look_Pos = Look_Pos + 1
            Token = "<"
            Exit Sub
         End If

       Case ">"

         If Mid$(ExprLine, Look_Pos + 1, 1) = "=" Then
            'Greator than or equal to
            Look_Pos = Look_Pos + 2
            Token = Chr$(GE)
            Exit Sub
          ElseIf Mid$(ExprLine, Look_Pos + 1, 1) = ">" Then
            Look_Pos = Look_Pos + 2
            Token = Chr$(shr)
            Exit Sub
          Else
            'Greator than
            Look_Pos = Look_Pos + 1
            Token = ">"
            Exit Sub
         End If

         tok_type = DELIMITER
      End Select

   End If

   If IsDelim(Mid$(ExprLine, Look_Pos, 1)) Then
      'Check if we have a Delimiter ;,+-<>^=(*)/\%
      Token = Token + Mid$(ExprLine, Look_Pos, 1) 'Get next char
      Look_Pos = Look_Pos + 1 'INC
      tok_type = DELIMITER 'Delimiter Token type
    ElseIf isDigit(Mid$(ExprLine, Look_Pos, 1)) Then
      'See if we are dealing with a Hexadecimal Value

      If Mid$(ExprLine, Look_Pos + 1, 1) = "x" Then

         Do While (isAlphaNum(Mid$(ExprLine, Look_Pos, 1)))
            Token = Token + Mid$(ExprLine, Look_Pos, 1)
            Look_Pos = Look_Pos + 1
            tok_type = HEXDIGIT
         Loop

         Exit Sub
      End If

      'Check if we are dealing with only digits 0 .. 9

      Do While (IsDelim(Mid$(ExprLine, Look_Pos, 1))) = 0
         Token = Token + Mid$(ExprLine, Look_Pos, 1) 'Get next char
         Look_Pos = Look_Pos + 1 'INC
      Loop

      tok_type = DIGIT 'Digit token type

    ElseIf isAlpha(Mid$(ExprLine, Look_Pos, 1)) Then
      'Check if we have strings Note no string support in this version
      ' this is only used for variables.

      Do While Not IsDelim(Mid$(ExprLine, Look_Pos, 1))
         Token = Token + Mid$(ExprLine, Look_Pos, 1)
         Look_Pos = Look_Pos + 1 'INC
         'tok_type = VARIABLE
         tok_type = LSTRING 'String token type
      Loop
    Else
      Abort 4, Mid$(ExprLine, Look_Pos, 1)
      tok_type = FINISHED
   End If

   If tok_type = LSTRING Then
      'check for identifiers

      If isIdent(Token) Then

         Select Case UCase(Token)
          Case "AND"
            Token = Chr(cAND)
            Exit Sub

          Case "OR"
            Token = Chr(cOR)
            Exit Sub

          Case "NOT"
            Token = "!"
            Exit Sub

          Case "IMP"
            Token = Chr(cIMP)
            Exit Sub

          Case "EQV"
            Token = Chr(cEqv)
            Exit Sub

          Case "DIV"
            Token = "\"
            Exit Sub

          Case "MOD"
            Token = "%"
            Exit Sub

          Case "XOR"
            Token = Chr(cXor)
            Exit Sub

          Case "SHL"
            Token = Chr(shl)
            Exit Sub

          Case "SHR"
            Token = Chr(shr)
            Exit Sub
         End Select

         tok_type = DELIMITER
         Exit Sub

       ElseIf IsIdentFunc(Token) Then
         tok_type = IDENTIFIER
         ' GetToken
         Exit Sub
       Else
         tok_type = VARIABLE
         Exit Sub
      End If

   End If

End Sub

Private Function GetVarData(Name As String) As Variant

   'Return data from a variable stored in the variable stack
   GetVarData = lVars(FindVarIdx(Name)).vValue

End Function

Public Sub Init()

   lVarCount = 0
   Erase lVars

End Sub

Private Function isAlpha(c As String) As Boolean

   'Return true if we only have letters a-z  A-Z
   isAlpha = UCase$(c) >= "A" And UCase$(c) <= "Z"

End Function

Private Function isAlphaNum(c As String) As Boolean

   isAlphaNum = (isDigit(c) Or isAlpha(c))

End Function

Private Function IsDelim(c As String) As Boolean

   'Return true if we have a Delimiter
   If InStr(" ;,+-<>^=(*)/\%&|!", c) Then IsDelim = True

End Function

Private Function isDigit(c As String) As Boolean

   'Return true when we only have a digit
   isDigit = (c >= "0") And (c <= "9")

End Function

Private Function isHex(HexVal As String) As Boolean

  Dim x As Integer
  Dim c As String

   For x = 1 To Len(HexVal)
      c = Mid$(HexVal, x, 1)

      Select Case UCase$(c)
       Case 0 To 9: isHex = True
       Case "A", "B", "C", "D", "E", "F": isHex = True
       Case Else
         isHex = False
         Exit For
      End Select

   Next x

End Function

Private Function isIdent(sIdentName As String) As Boolean

  Dim x As Integer
  Dim Idents As Variant

   Idents = Split(Str_Ops, ",")

   For x = 0 To UBound(Idents)

      If LCase$(Idents(x)) = LCase$(sIdentName) Then
         isIdent = True
         Exit For
      End If

   Next x

   x = 0
   Erase Idents

End Function

Private Function IsIdentFunc(sIdentName As String) As Boolean

  Dim x As Integer
  Dim Idents As Variant

   Idents = Split(Str_Funcs, ",")

   For x = 0 To UBound(Idents)

      If LCase$(Idents(x)) = LCase$(sIdentName) Then
         IsIdentFunc = True
         Exit For
      End If

   Next x

   x = 0
   Erase Idents

End Function

Private Function isWhite(c As String) As Boolean

   'Return true if we find a white space
   isWhite = (c = " ") Or (c = vbTab)

End Function

Private Sub PushBack()

  Dim tok_len As Integer

   tok_len = Len(Token)
   Look_Pos = Look_Pos - tok_len

End Sub

Public Sub SetVar(vIdx, Optional lData As Variant = 0)

   'Set a variables value, by using the variables index vIdx
   lVars(vIdx).vValue = lData

End Sub

Public Sub SetVarName(sName, Optional lData As Variant = 0)

  Dim vIdx As Long
  Dim lFound As Boolean

   'Set a variables value, by using the variables index vIdx
   lFound = False
   For vIdx = 0 To UBound(lVars)
       If (sName = lVars(vIdx).vName) Then
           lFound = True
           Exit For
       End If
       
   Next vIdx
   
   If (lFound = True) Then
       lVars(vIdx).vValue = lData
   End If

End Sub
