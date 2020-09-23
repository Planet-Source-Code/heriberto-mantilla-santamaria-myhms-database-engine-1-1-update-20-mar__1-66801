Attribute VB_Name = "Eingana_Encryption"
' Eingana Encryption Bas.
' This is my first attempt at encryption
' Please let me know what you think
' If you find this easy to break please
' e-mail me with suggestions to improve
' my e-mail is SteveRuben@Gmail.com
' Thank You - Steve Ruben

Option Explicit

 Public Key    As String
 Public KeyTwo As String
 Public Key1   As Byte
 Public Key2   As String
 Public Key3   As Byte
 
 Private Differance As Variant
 
Public Function KeySetup(ByVal Key As String) As Boolean
  Dim KeyLen As Integer
  Dim Ascii As Integer
  Dim AsciiString As String
  Dim AsciiConversion As Double
  Dim i As Integer
  Dim First As Byte
  Dim Last As Byte
  KeyLen = Len(Key)
  If KeyLen > 128 Then
    KeySetup = False
    Exit Function
  End If
  If KeyLen < 8 Then
    KeySetup = False
    Exit Function
  End If
  For i = 1 To KeyLen
    Ascii = Ascii + Asc(Mid(Key, i, 1))
  Next i
  Key1 = Ascii \ i
  For i = 1 To KeyLen
    AsciiString = AsciiString & Asc(Mid(Key, i, 1))
  Next i
  Key2 = AsciiString
  First = Asc(Mid(Key, 1, 1))
  Last = Asc(Mid(Key, KeyLen, 1))
  Key3 = (First + Last) \ 2
  KeySetup = True
End Function

Public Function EncryptKey2(ByVal Text As String) As String
  Dim TextLen As Double
  Dim Counter As Double
  Dim Ascii As Byte
  Dim kAscii As Byte
  Dim CodedByte As Byte
  Dim Output As String
  Dim i As Double

  TextLen = Len(Text)
  Counter = 1
  For i = 1 To TextLen
    If (Counter > Len(Key2)) Then
      Counter = 1
    End If
    Ascii = Asc(Mid(Text, i, 1))
    kAscii = Val(Mid(Key2, Counter, 1))
    CodedByte = Ascii + kAscii
    Output = Output & Chr(CodedByte)
    Counter = Counter + 1
  Next i
  EncryptKey2 = Output
End Function

Public Function DecryptKey2(ByVal Text As String) As String
  Dim TextLen As Double
  Dim Counter As Double
  Dim Ascii As Byte
  Dim kAscii As Byte
  Dim CodedByte As Byte
  Dim Output As String
  Dim i As Double

  TextLen = Len(Text)
  Counter = 1
  For i = 1 To TextLen
    If (Counter > Len(Key2)) Then
      Counter = 1
    End If
    Ascii = Asc(Mid(Text, i, 1))
    kAscii = Val(Mid(Key2, Counter, 1))
    CodedByte = Ascii - kAscii
    Output = Output & Chr$(CodedByte)
    Counter = Counter + 1
  Next i
  DecryptKey2 = Output
End Function

Public Function TextShiftForward(ByVal Text As String) As String
  Dim TextLen As Double
  Dim i As Double
 
  TextLen = Len(Text)
  For i = 1 To Key1
    Text = Mid$(Text, 2, TextLen) & Mid$(Text, 1, 1)
  Next i
  TextShiftForward = Text
End Function

Public Function TextShiftReverse(ByVal Text As String) As String
  Dim TextLen As Double
  Dim i As Double

  TextLen = Len(Text)
  For i = 1 To Key1
    Text = Mid$(Text, TextLen, 1) & Mid$(Text, 1, TextLen - 1)
  Next i
  TextShiftReverse = Text
End Function

Public Function EncryptKey3(ByVal Text As String) As String
  Dim TextLen As Double
  Dim i As Double
  Dim TextArray() As Byte
  Dim KeyArray() As Byte
  Dim KeyLen As Integer
  Dim Key2Len As Integer
  Dim LargeKey As String
  Dim k1Start As Byte
  Dim k2Start As Byte
  Dim TempKey As String
  Dim kAddition As Integer
  Dim kDivision As Integer
  Dim LargeKeyLen As Integer
  Dim Counter As Double
  Dim NewLetter As String
  Dim FinalString As String
  Dim NewNum As Byte

  KeyLen = Len(Key)
  Key2Len = Len(KeyTwo)
  TextLen = Len(Text)
  If KeyLen > Key2Len Then
    Differance = KeyLen - Key2Len
On Error Resume Next
    KeyTwo = KeyTwo & Space(Differance)
  Else
    Differance = Key2Len = KeyLen
On Error Resume Next
    Key = Key & Space(Differance)
  End If
  KeyLen = Len(Key)
  Key2Len = Len(KeyTwo)
  For i = 1 To KeyLen
    k1Start = Asc(Mid$(Key, i, 1))
    k2Start = Asc(Mid$(KeyTwo, i, 1))
    kAddition = k1Start + k2Start
    kDivision = kAddition \ Sqr(Key3)
    kAddition = k1Start Mod k2Start
    TempKey = TempKey & Chr$(kDivision + kAddition)
  Next i
  ReDim TextArray(1 To TextLen)
  ReDim KeyArray(1 To TextLen)
  For i = 1 To TextLen
    TextArray(i) = Asc(Mid$(Text, i, 1))
  Next i
  For i = 1 To KeyLen
    LargeKey = LargeKey & Mid$(Key, i, 1) & Mid$(KeyTwo, i, 1) & Mid$(TempKey, i, 1)
  Next i
  LargeKeyLen = Len(LargeKey)
  Counter = 1
  For i = 1 To TextLen
    If (Counter > LargeKeyLen) Then
      Counter = 1
    End If
    KeyArray(i) = Asc(Mid$(LargeKey, Counter, 1))
    Counter = Counter + 1
  Next i
  Counter = 1
  For i = 1 To TextLen
    If Counter > LargeKeyLen Then Counter = 1
    NewNum = Int(Sqr(KeyArray(i) + KeyArray(Counter)))
    NewLetter = Chr(TextArray(i) + NewNum)
    FinalString = FinalString & NewLetter
    Counter = Counter + 1
  Next i
  EncryptKey3 = FinalString
End Function

Public Function DecryptKey3(ByVal Text As String) As String
  Dim TextLen As Double
  Dim i As Double
  Dim TextArray() As Byte
  Dim KeyArray() As Byte
  Dim KeyLen As Integer
  Dim Key2Len As Integer
  Dim LargeKey As String
  Dim k1Start As Byte
  Dim k2Start As Byte
  Dim TempKey As String
  Dim kAddition As Integer
  Dim kDivision As Integer
  Dim LargeKeyLen As Integer
  Dim Counter As Double
  Dim NewLetter As String
  Dim FinalString As String
  Dim NewNum As Byte
 
  KeyLen = Len(Key)
  Key2Len = Len(KeyTwo)
  TextLen = Len(Text)
  If KeyLen > Key2Len Then
    Differance = KeyLen - Key2Len
On Error Resume Next
    KeyTwo = KeyTwo & Space(Differance)
  Else
    Differance = Key2Len = KeyLen
On Error Resume Next
    Key = Key & Space(Differance)
  End If
  KeyLen = Len(Key)
  Key2Len = Len(KeyTwo)
  For i = 1 To KeyLen
    k1Start = Asc(Mid$(Key, i, 1))
    k2Start = Asc(Mid$(KeyTwo, i, 1))
    kAddition = k1Start + k2Start
    kDivision = kAddition \ Sqr(Key3)
    kAddition = k1Start Mod k2Start
    TempKey = TempKey & Chr$(kDivision + kAddition)
  Next i
  ReDim TextArray(1 To TextLen)
  ReDim KeyArray(1 To TextLen)
  For i = 1 To TextLen
    TextArray(i) = Asc(Mid$(Text, i, 1))
  Next i
  For i = 1 To KeyLen
    LargeKey = LargeKey & Mid$(Key, i, 1) & Mid$(KeyTwo, i, 1) & Mid$(TempKey, i, 1)
  Next i
  LargeKeyLen = Len(LargeKey)
  Counter = 1
  For i = 1 To TextLen
    If (Counter > LargeKeyLen) Then
      Counter = 1
    End If
    KeyArray(i) = Asc(Mid$(LargeKey, Counter, 1))
    Counter = Counter + 1
  Next i
  Counter = 1
  For i = 1 To TextLen
    If (Counter > LargeKeyLen) Then
      Counter = 1
    End If
    NewNum = Int(Sqr(KeyArray(i) + KeyArray(Counter)))
    NewLetter = Chr(TextArray(i) - NewNum)
    FinalString = FinalString & NewLetter
    Counter = Counter + 1
  Next i
  DecryptKey3 = FinalString
End Function
