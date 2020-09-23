Attribute VB_Name = "modDataFunc"
'*******************************************************'
'* Program:       myHMS Database Engine                *'
'*******************************************************'
'* Author:        Heriberto Mantilla Santamaría        *'
'*******************************************************'
'* Build DataBase using Plan Text File.                *'
'*                                                     *'
'* Based on Jim's original code [CodeId=60559] allows  *'
'* to create a Database using plane files, you can     *'
'* create:                                             *'
'*                                                     *'
'* 1. Several tables.                                  *'
'* 2. Can put create an username and password for the  *'
'*    database.                                        *'
'* 3. Add fields in any table specifically.            *'
'* 4. Select certain fields in a table.                *'
'*                                                     *'
'======================================================*'
'* In development:                                     *'
'*                                                     *'
'* 1. To put to carry out searches.                    *'
'* 2. To upgrade a chart according to parameters.      *'
'* 3. To compress the generated files.                 *'
'*-----------------------------------------------------*'
'* CREDITS AND THANKS                                  *'
'*-----------------------------------------------------*'
'*           Build DataBase In Text File               *'
'*             Jim Jose [CodeId=60559]                 *'
'======================================================*'
'*      Binary File Password Encryptor class!!!        *'
'*              Cahaltech [CodeId=36457]               *'
'*-----------------------------------------------------*'
'* Dedicated to                                        *'
'*-----------------------------------------------------*'
'* Ivan T.P (halo sahabatku apa kabar)                 *'
'* mibi (saya mencintai kamu)                          *'
'* sakura_tsukino (my little sister)                   *'
'* Gaby (my best girl friend)                          *'
'* Pancho funes                                        *'
'*******************************************************'
'*                   Version 1.0.0                     *'
'*******************************************************'
'*                                                     *'
'* Note:     Comments, suggestions, doubts or bug      *'
'*           reports are wellcome to these e-mail      *'
'*           addresses:                                *'
'*                                                     *'
'*                  heri_05-hms@mixmail.com or         *'
'*                  hcammus@hotmail.com                *'
'*                                                     *'
'*        Please rate my work on this control.         *'
'*    That lives the Soccer and the América of Cali    *'
'*             Of Colombia for the world.              *'
'*******************************************************'
'*        All rights Reserved © HACKPRO TM 2006        *'
'*******************************************************'
Option Explicit

Public HMSEngine As myHMSEngine

Private InCent As Integer ' Error value.

Private Const Alphabethic = "[A-Z a-z 0-9 _]"
Private Const DataType = _
    "BINARY|CHAR|INTEGER|DATETIME|LONG|LONGTEXT|VARCHAR|DOUBLE|SMALLINT|BYTE|CURRENCY|BOOL|"
Private Const Structure = "CREATE TABLE|CREATE DATABASE|CREATE USER|CREATE PASSWORD|TABLE" & _
    " NAME|TABLE|"


Public Function CompareData(ByVal iData As String, _
                            ByVal Value As String, _
                            ByVal PKey As String, _
                            ByVal NotNull As String) As Boolean

  Dim TypeData As Variant
  Dim  tData As Variant
  Dim ParOpen  As Integer
  Dim ParClose As Integer

    On Error GoTo mError
    InCent = -1
    ParOpen = InStr(1, iData, "(")
    ParClose = InStr(1, iData, ")")

    If ((CInt(PKey) = 1) Or (NotNull = "N")) And (Value = "") Then
        InCent = 0
        CompareData = True
        Exit Function
    End If

    Select Case iData
    Case "INTEGER"
        InCent = 1
        TypeData = CInt(Value)

    Case "BYTE"
        InCent = 2
        TypeData = CByte(Value)

    Case "DOUBLE"
        InCent = 3
        TypeData = CDbl(Value)

    Case "DATETIME"
        InCent = 4
        TypeData = CDate(Value)

    Case "LONG"
        InCent = 5
        TypeData = CLng(Value)

    Case "SMALLINT"
        InCent = 6
        TypeData = CDec(Value)

    Case "CURRENCY"
        InCent = 7
        TypeData = CCur(Value)

    Case "BOOL"
        InCent = 8
        TypeData = CBool(Value)

    Case Else

        If (ParOpen > 0) And (ParClose > 0) Then
            ' See if is VARCHAR or CHAR data type.
            tData = CStr(UCase$(Mid$(iData, 1, ParOpen - 1)))

            If (tData = "CHAR") Or (tData = "VARCHAR") Then
                tData = CInt(Mid$(iData, ParOpen + 1, (ParClose - 1) - ParOpen))

                If (Len(Value) <= tData) Then
                    CompareData = False
                Else
                    InCent = 9
                    CompareData = True
                End If

            End If
        End If
    End Select
    CompareData = False
    Exit Function
    mError:
  Dim lMessage As String

    CompareData = True

    Select Case InCent
    Case 0
        lMessage = "The Primary Key is NULL"

    Case 1 To 9
        lMessage = "ERROR of conversion Data type"

    Case Else
        lMessage = ""
    End Select

    If (lMessage <> "") Then
        MsgBox lMessage, vbCritical + vbOKOnly, "myHMS Database"
    End If

End Function

Public Function CountCharacters(ByVal TheString As String, ByVal CharToCheck As String) As Integer

  Dim mPos As Long
  Dim  ReturnAgain As Boolean
  Dim Char As String

    ' Count the number of occurrences of one string within another string.
    CountCharacters = 0

    For mPos = 1 To Len(TheString)

        If (mPos < (Len(TheString) + 1 - Len(CharToCheck))) Then
            Char = Mid$(TheString, mPos, Len(CharToCheck))
            ReturnAgain = True
        Else
            Char = Mid$(TheString, mPos)
            ReturnAgain = False
        End If

        If (Char = CharToCheck) Then CountCharacters = CountCharacters + 1
        If (ReturnAgain = False) Then Exit For
    Next mPos

End Function

Public Function FindStructure(ByVal iData As String) As String

  Dim tData As Variant
  Dim  lData As String
  Dim iPos  As Integer

    ' Find the principal structure database.
    tData = Split(Structure, "|")
    FindStructure = ""

    For iPos = 0 To UBound(tData)
        lData = Mid$(iData, 1, Len(tData(iPos)))

        If (lData = tData(iPos)) Then
            FindStructure = lData
            Exit For
        End If

    Next iPos

End Function

Public Function InDataType(ByVal iData As String) As Boolean

  Dim rData As Variant
  Dim  ParOpen  As Integer
  Dim  tData As String
  Dim iPos  As Integer
  Dim  ParClose As Integer

    ' Search if the data type is valid.
    InDataType = False
    rData = Split(DataType, "|")

    For iPos = 0 To UBound(rData)
        ParOpen = InStr(1, iData, "(")
        ParClose = InStr(1, iData, ")")

        If (rData(iPos) = UCase$(iData)) Then
            InDataType = True
            Exit For
        ElseIf (ParOpen > 0) And (ParClose > 0) Then
            ' See if is VARCHAR or CHAR data type.
            tData = UCase$(Mid$(iData, 1, ParOpen - 1))

            If (tData = "CHAR") Or (tData = "VARCHAR") Then
                tData = Mid$(iData, ParOpen + 1, (ParClose - 1) - ParOpen)

                If (IsNumeric(tData) = True) Then
                    If (CInt(tData) <= 256) And (CInt(tData) > 1) Then
                        InDataType = True
                        Exit For
                    End If

                End If
            End If
        End If
    Next iPos

End Function

Public Function RemoveComment(ByVal pLine As String, Optional ByVal Token As String = "'") As String

  Dim pComa As Long
  Dim  nCarac    As String
  Dim  initPos As Long
  Dim AllOk As Boolean
  Dim  pCount As Long

    ' Remove the comments of a string.
    pComa = -1
    initPos = 1
    AllOk = False

    Do While (AllOk = False) And (Len(pLine) > 0)
        ' Search the position of the simple quotation marks.
        pComa = InStr(initPos, pLine, Token)
        If (pComa = 0) Then Exit Do
        ' We take the text until the position of the simple quotation marks.
        nCarac = RTrim$(Mid$(pLine, 1, pComa))
        pCount = CountCharacters(nCarac, Chr$(34)) Mod 2

        If (pCount = 1) Then
            initPos = pComa + 1
            AllOk = False
        Else
            AllOk = True
            Exit Do
        End If

    Loop
    If (AllOk = True) Then '* Return the string without comment.
        RemoveComment = Mid$(pLine, 1, pComa - 1)
    Else
        RemoveComment = pLine
    End If

End Function

Public Function ValidToken(ByVal isValue As String) As Boolean

  Dim iPos As Integer
  Dim  Carac As String

    ValidToken = True
    isValue = Trim$(isValue)
    Carac = Mid$(isValue, 1, 1)

    If (Carac = "_") Or (isValue = "") Or (IsNumeric(Carac) = True) Or (Len(isValue) > 10) Then
        ValidToken = True
        Exit Function
    End If

    For iPos = 1 To Len(isValue)
        Carac = Mid$(isValue, iPos, 1)

        If (Carac = "") Then
            ValidToken = False
            Exit For
        ElseIf Not (Carac Like Alphabethic) Then
            ValidToken = False
            Exit For
        End If

    Next iPos

End Function

