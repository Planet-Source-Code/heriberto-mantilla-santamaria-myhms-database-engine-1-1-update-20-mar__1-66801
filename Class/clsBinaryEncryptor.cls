VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBinaryEncryptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* Developed by Cahaltech
Option Explicit

Private Declare Function GetTempPath Lib "kernel32" _
        Alias "GetTempPathA" ( _
        ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long

Public Function DecryptFile(Source As String, Password As String) As String

On Error GoTo 1
  Dim ByteIn() As Byte
  Dim ByteOut() As Byte
  Dim i As Long
  Dim j As Long
  Dim PL As Integer
  Dim ChrBNow As Integer
  Dim PosNow As Integer
  Dim TempDByte As Integer
  Dim ByteFinal As Integer
  Dim FleeFile  As Integer
  Dim Bit() As Byte

    ByteIn() = ReadBinaryArray(Source)
    Bit() = ReadBinaryArray(Source)

    ReDim ByteOut(LBound(ByteIn) To UBound(ByteIn)) As Byte

    PL = Len(Password)

    For i = LBound(ByteIn) To UBound(ByteIn)

        PosNow = i Mod PL

        ChrBNow = AscB(Mid(Password, PosNow + 1, 1)) Xor 17

        TempDByte = (ByteIn(i) - ChrBNow)

        If TempDByte < 0 Then

            ByteFinal = 256 - Abs(TempDByte)

        Else

            ByteFinal = TempDByte

        End If

        ByteOut(i) = ByteFinal

        If i Mod 500 = 0 Then
            DoEvents
            'RaiseEvent DecryptProgress(i, UBound(ByteIn))
        End If

    Next i

    SaveBinaryArray TempPathName & "myHMS.tmp", ByteOut
  
    FleeFile = FreeFile
    Open TempPathName & "myHMS.tmp" For Binary As FleeFile
        ' Getting the DataBase Parameters.
        Get FleeFile, 1, Bit()
        DecryptFile = StrConv(Bit, vbUnicode)
    Close FleeFile
    
    Kill TempPathName & "myHMS.tmp"
    Exit Function
1
    DecryptFile = ""

End Function

Public Function EncryptFile(Source As String, Password As String) As Boolean

On Error GoTo 1
  Dim ByteIn() As Byte
  Dim ByteOut() As Byte
  Dim ChrBNow As Integer
  Dim PosNow As Integer
  Dim TempByte As Integer

  ' Decrypt
  Dim TempDByte As Integer
  Dim ByteFinal As Integer
  Dim i As Long
  Dim j As Long
  Dim PL As Integer

    ByteIn() = ReadBinaryArray(Source)

    ReDim ByteOut(LBound(ByteIn) To UBound(ByteIn)) As Byte

    PL = Len(Password)

    For i = LBound(ByteIn) To UBound(ByteIn)

        PosNow = i Mod PL

        ChrBNow = AscB(Mid$(Password, PosNow + 1, 1)) Xor 17
        'ChrBNow = 1

        TempByte = (ByteIn(i) + ChrBNow) Mod 256

        TempDByte = (TempByte - ChrBNow)

        If TempDByte < 0 Then

            ByteFinal = 256 - Abs(TempDByte)

        Else

            ByteFinal = TempDByte

        End If

        ByteOut(i) = TempByte
        
        If i Mod 500 = 0 Then
            DoEvents
        End If

    Next i

    SaveBinaryArray TempPathName & "myHMS.tmp", ByteOut
    FileCopy TempPathName & "myHMS.tmp", Source
    Kill TempPathName & "myHMS.tmp"
    EncryptFile = True
    Exit Function
    
1
    EncryptFile = False

End Function

Public Function ReadBinaryArray(ByVal Source As String)

  Dim bytBuf() As Byte
  Dim intN     As Long
  Dim t        As Integer
  Dim n As Long

On Error GoTo myErr
    t = FreeFile
    Open Source For Binary Access Read As #t
        ReDim bytBuf(1 To LOF(t)) As Byte
        Get #t, , bytBuf()

        ReadBinaryArray = bytBuf()
myErr:
    Close #t

End Function

Public Sub SaveBinaryArray(ByVal Filename As String, WriteData() As Byte)

  Dim t As Integer

On Error GoTo myErr
    t = FreeFile
    Open Filename For Binary Access Write As #t
        Put #t, , WriteData()
myErr:
    Close #t

End Sub

Private Function TempPathName() As String

  Dim strTemp As String

    ' Returns the name of the temporary directory of Windows.
    strTemp = String$(100, Chr$(0)) '* Create a buffer.
    GetTempPath 100, strTemp  ' Get the temporary path.
    ' Strip the rest of the buffer.
    TempPathName = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)

End Function
