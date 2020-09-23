Attribute VB_Name = "modFile"
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

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
        Alias "GetOpenFileNameA" ( _
        pOpenfilename As OpenFileName) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
        Alias "GetSaveFileNameA" ( _
        pOpenfilename As OpenFileName) As Long

Private Type OpenFileName
    lStructSize       As Long
    hWndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Private OFName       As OpenFileName
Private DEncrypt     As clsBinaryEncryptor

Public xOpenFile     As Boolean


Public Function AppDir() As String

    If (Right$(App.Path, 1) <> "\") Then
        AppDir = App.Path & "\"
    Else
        AppDir = App.Path
    End If

End Function

Public Sub CreateFile(ByVal NameFile As String, ByVal Text As String)

  Dim FleeFile As Integer

    FleeFile = FreeFile()
    Open NameFile For Binary Access Write As #FleeFile
    Put #FleeFile, , Text
    Close #FleeFile
    EncryptFile NameFile

End Sub

Public Function Decrypt(ByVal NameFile As String) As String

    Set DEncrypt = New clsBinaryEncryptor
    Decrypt = DEncrypt.DecryptFile(NameFile, "myHMS Database by HACKPRO TM")

End Function

Public Sub EncryptFile(ByVal NameFile As String)

    Set DEncrypt = New clsBinaryEncryptor
    DEncrypt.EncryptFile NameFile, "myHMS Database by HACKPRO TM"

End Sub

Public Function FileExits(ByVal Exists As String) As Boolean

  Dim FindFolder As String

    FileExits = False
    On Error GoTo FileError
    FindFolder = Dir(Exists)

    If (FindFolder <> "") Then
        FileExits = True  '* Found the file or folder.
        Exit Function
    End If

    FileError:
    FileExits = False

End Function

Private Sub IniFile()

    OFName.lStructSize = Len(OFName)
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Database (*.dms)" + Chr$(0) + "*.dms"
    OFName.lpstrDefExt = ".dms"
    OFName.lpstrTitle = "Open database"
    OFName.nFilterIndex = 4
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.flags = &H80000 + &H400 + &H1000 + &H4 + &H8 + &H2 + &H800
    OFName.lpstrInitialDir = AppDir & "data"

End Sub

Public Function OpenFile(ByVal NameFile As String, _
                         Optional ByVal RemoveComments As Boolean = False) As String

  Dim FleeFile  As Integer
  Dim Bit(1000) As Byte
  Dim m_DataID  As String

    If (FileExits(NameFile) = True) Then
        OpenFile = Decrypt(NameFile)
        xOpenFile = True
        Exit Function
    Else
        xOpenFile = False
    End If

End Function

Public Function ShowOpen(ByVal hWnd As Long) As String

  Dim iPos As Integer
  Dim  iText As String

    OFName.hWndOwner = hWnd
    Call IniFile
    ShowOpen = ""

    If (GetOpenFileName(OFName) <> 0) Then
        iText = Trim$(OFName.lpstrFile)
        ShowOpen = Mid$(iText, 1, Len(iText) - 1)
        iPos = InStrRev(ShowOpen, "\")
        Call ChDir(Mid$(ShowOpen, 1, Abs(iPos - 1)))
    Else
        ShowOpen = ""
    End If

End Function

