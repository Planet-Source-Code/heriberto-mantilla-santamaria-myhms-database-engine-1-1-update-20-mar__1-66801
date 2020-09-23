VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "myHMSEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private FileOp          As String
Private iPos            As Long
Private xSplit          As Variant
Private TempDt          As String
Private tPos            As Long
Private ySplit          As Variant
Private mTFieldCount    As Integer
Private mRecordCount    As Long
Private mActualPos      As Long
Private m_TableName     As String

Private bOpenDatabase   As Boolean
Private bOpenRecordset  As Boolean
Private bOpenTable      As Boolean

Private NTables()       As String
Private lRecord()       As String

Private m_sDatabaseName As String  ' Database name.
Private m_sUserName     As String  ' User of database.
Private m_sPassword     As String  ' Password of database.

Private sc, m
 

Private Property Get ActualPos() As Long

    ActualPos = mActualPos

End Property

Private Sub Class_Initialize()

    mActualPos = 0

End Sub


Public Sub ChangedDatabaseName(ByVal Source As String, ByVal DatabaseName As String, Optional ByVal Username As String = "", _
                             Optional ByVal Password As String = "")
    
   Dim xData As String
   Dim lPos As Long
   Dim iCell As Long
   Dim TempVal As String
    
    If (OpenDatabase(Source, Username, Password) = True) Then
        xData = ""
        xData = BFinalToken & vbCrLf & CommentToken & "myHMS Database" & vbCrLf & BFinalToken
        xData = xData & vbCrLf & vbCrLf & CommentSchema & Trim$(DatabaseName) & vbCrLf & CommentToken & _
            vbCrLf & vbCrLf & "CREATE DATABASE `" & Trim$(DatabaseName) & "`;"
        xData = xData & vbCrLf & vbCrLf & CommentToken & vbCrLf & CommentToken & "CREATE USER &" & _
            " PASSWORD" & vbCrLf & CommentToken & vbCrLf & vbCrLf & "CREATE USER `" & _
            Trim$(Username) & "`;" & vbCrLf & "CREATE PASSWORD `" & Trim$(Password) & "`;"
    
        ' Create tables in principal file.
        xData = xData & vbCrLf & vbCrLf & CommentToken & vbCrLf & CommentToken & "CREATE TABLES" & _
            vbCrLf & CommentToken & vbCrLf & vbCrLf

        For iCell = 0 To (TotalTables - 1)

            If (iCell < (TotalTables - 1)) Then
                TempVal = vbCrLf
            Else
                TempVal = ""
            End If

            xData = xData & "TABLE NAME `" & TableNames(iCell) & "`;" & TempVal
        Next iCell

        ' Create principal file.
        lPos = InStrRev(Source, "\")
        
        Source = Mid$(Source, 1, lPos)
        CreateFile Source & Trim$(DatabaseName) & ".dms", xData
    Else
        Err.Raise 524, "ChangedDatabaseName", "Can't changed the name of this database."
    End If
    
End Sub


'-------------------------------------------------------------------------
' Procedure  : CloseDatabase
' Author     : Heriberto Mantilla Santamaría
' Input      : None
' OutPut     : None
' Purpose    : Close the Database opened
'-------------------------------------------------------------------------
Public Sub CloseDatabase()

    bOpenDatabase = False
    DatabaseName = ""
    Username = ""
    Password = ""
    Erase TableNames
    CloseRecordset
    CloseTable

End Sub


'-------------------------------------------------------------------------
' Procedure  : CloseRecordset
' Author     : Heriberto Mantilla Santamaría
' Input      : None
' OutPut     : None
' Purpose    : Close the Recordset opened
'-------------------------------------------------------------------------
Public Sub CloseRecordset()

    mRecordCount = -1
    bOpenRecordset = False
    Erase lRecord
    m_TableName = ""

End Sub


'-------------------------------------------------------------------------
' Procedure  : CloseTable
' Author     : Heriberto Mantilla Santamaría
' Input      : None
' OutPut     : None
' Purpose    : Close the table opened
'-------------------------------------------------------------------------
Public Sub CloseTable()

    Erase NTables
    bOpenTable = False

End Sub


Public Property Get DatabaseName() As String

    DatabaseName = m_sDatabaseName

End Property

Public Property Let DatabaseName(ByVal sDatabaseName As String)

    m_sDatabaseName = sDatabaseName

End Property


Public Property Get DatabaseIsOpen() As Boolean
    
    DatabaseIsOpen = bOpenDatabase
    
End Property


'-------------------------------------------------------------------------
' Procedure  : Delete
' Author     : Heriberto Mantilla Santamaría
' Input      : Table, Columns, Values()
' OutPut     : None
' Purpose    : To delete the current record
'-------------------------------------------------------------------------
Public Sub Delete(ByVal Table As String, ByVal Columns As String, ByRef Values() As String)
    
    If (OpenSTable(Table, IIf(Columns = "*", "", Columns)) = True) Then
        
    Else
        Err.Raise 520, "Insert", "Cannot insert the fields."
    End If
    
End Sub


'-------------------------------------------------------------------------
' Procedure  : DeleteAll
' Author     : Heriberto Mantilla Santamaría
' Input      : TableName
' OutPut     : None
' Purpose    : To delete the current record
'-------------------------------------------------------------------------
Public Sub DeleteAll(ByVal TableName As String)

  Dim zsData As String
  Dim NameFile As String
  Dim FleeFile As Integer

    If (TableInDatabase(TableName) = False) Then
        Err.Raise 518, "OpenRecordSet", "Cannot open Recordset."
    Else
        NameFile = AppDir & "data\" & DatabaseName & "\" & TableName & ".udt"
        FleeFile = FreeFile
        Open NameFile For Binary Access Write As FleeFile
            Put FleeFile, 1, ""
        Close FleeFile
        EncryptFile NameFile
    End If
    
End Sub


Private Function EvalSQL(ByVal SQL As String) As Boolean
    
    Set sc = CreateObject("ScriptControl")
    sc.Language = "VBScript"
    Set m = sc.Modules.Add("modEval")

On Error GoTo myErr:
    EvalSQL = False
    If (m.Eval(SQL) = True) Then
        EvalSQL = True
    End If
    
    Exit Function

myErr:

    EvalSQL = False
    
End Function

Private Sub ExtractLiteral(ByRef Where_T As String)
    
  Dim iLen As Long
  Dim iPos As Long
  Dim lTemp As String
  Dim sSandw As String
  
  iLen = Len(Where_T)
  lTemp = ""
  For iPos = 1 To iLen
      If (Mid$(Where_T, iPos, 1) = "'") Then
          sSandw = Trim$(FuncInstrSandwich(iPos, Where_T, "'", "'", True))
          iPos = LendPos
          lTemp = lTemp & "0"
          If (iPos = 0) Then
              isAbort = True
              Exit For
          End If
          
      Else
          lTemp = lTemp & Mid$(Where_T, iPos, 1)
      End If
      
  Next iPos
  
  Where_T = lTemp
    
End Sub


'-------------------------------------------------------------------------
' Property   : Fields
' Author     : Heriberto Mantilla Santamaría
' Input      : Index
' OutPut     : Specific field of table opened
' Purpose    : Get field of specific table opened
'-------------------------------------------------------------------------
Public Property Get Fields(ByVal Index As Integer) As String

    On Error GoTo myErr

    If (bOpenRecordset = True) Then
        Fields = lRecord(ActualPos, Index)
    Else
        Err.Raise 511, "Fields", "Error opened the Field."
    End If

    Exit Property
myErr:
    Fields = ""

End Property


'-------------------------------------------------------------------------
' Procedure  : Find
' Author     : Heriberto Mantilla Santamaría
' Input      : Table, Columns, Values()
' OutPut     : None
' Purpose    : Find one record
'-------------------------------------------------------------------------
Public Sub Find(ByVal Table As String, ByVal Columns As String, ByRef Values() As String)
    
    If (OpenSTable(Table, Columns) = True) Then
        
    Else
        Err.Raise 520, "Insert", "Cannot insert the fields."
    End If
    
End Sub


Private Function FindWhere(ByVal Where As String, ByVal Item As String, ByVal Value As String) As String

  Dim sSandw As String
  Dim lPos As Long
  Dim lLen As Long

  lLen = Len(Where)
  sSandw = ""
  For lPos = 1 To lLen
      If (Mid$(Where, lPos, 1) = " ") Then
          sSandw = Replace$(sSandw, Item, Value)
          sSandw = sSandw & " "
      Else
          sSandw = sSandw & Mid$(Where, lPos, 1)
      End If
      
  Next lPos
  
  FindWhere = sSandw
  
End Function


'  This is a function that greatly expands the usefulness of the instr _
   function..this function looks for what is between 2 other strings _
   for example: _
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: _
 _
   lets say your looking to extract the host address of a full url. _
   In other words..what lies between "http//www." and the first "/" _
   followwing the http/www. _
   You could do an instr looking for the "http" part _
   then do an instr for the "/" part _
   and if they both return a nonzero then the word _
   were looking for starts at the first instr + the _
   len of the first word ("htp://www."-which is 10) _
   and ends at the second instr..so we would use the _
   mid() function. _
   This function does all this for you _
-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.- _
  PARAMETERS:
 '  lngStartPos:   [required]  :  The start point of the search for the first _
                                  word (strFindFirst$) same as instr _
    strToSearch$:  [required]  :  The string were doing this function on :-8 _
    strFindFirst$: [required]  :  The first string to search 4 in strToSearch$ _
    strFindEnd$:   [required]  :  The second string to search 4 instrToSearch$ _
    bCaseMatters   [optional]  :  Whether or not 2 take case into consideration _
    lngMaxSpreadLen[optional]  :  The maximum allowabled len between the 2 _
                                  strings (strFindFirst$ and strFindEnd$) _
-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-
Public Function FuncInstrSandwich(lngStartPos&, strToSearch$, strFindFirst$, strFindEnd$, Optional ByVal bCaseMatters As Boolean = False, Optional ByVal lngMaxSpreadLen As Long = -1) As String
 
  Dim l2Start As Long, l2 As Long, l1 As Long
 
    ' Set starting pos to avoid errors.
    If (lngStartPos < 1) Then lngStartPos = 1
    ' I case is irrelevant the lower _
        case everything.
    If (bCaseMatters = False) Then
        strToSearch$ = LCase$(strToSearch$)
        strFindFirst$ = LCase$(strFindFirst$)
        strFindEnd$ = LCase$(strFindEnd$)
    End If
 
    ' Find the starting pos of the first string.
    l1 = InStr(lngStartPos, strToSearch$, strFindFirst$)
    ' If it's found, search for the second part, _
        the start pos being return of the first _
        inst (l1) + len of that string.
    If (l1 <> 0) Then
        ' Property.
        l2Start = (Len(strFindFirst$) + l1)
        l2 = InStr(l2Start, strToSearch$, strFindEnd$)
        ' If the second string is found...
        If (l2 <> 0) Then
            ' This means user HAS NOT specified a max spead len or _
                he HAS specified a max spread len and the len of the _
                string between searchword1 and searchword2 <= lngMaxSpreadLen.
            If (lngMaxSpreadLen <= 0) Or _
                (lngMaxSpreadLen > 0) And _
                ((l2 - l2Start) <= lngMaxSpreadLen) _
            Then
                ' The actual string that lies between l2Start and l2.
                FuncInstrSandwich = Mid$(strToSearch$, l2Start, (l2 - l2Start))
                ' Return the start and end pos of the sandwich string.
                Let LStartPos = l2Start
                Let LendPos = l2
            Else
                FuncInstrSandwich = ""
            End If
                
        Else
            ' Second string not founds so return -1.
            FuncInstrSandwich = ""
        End If
    ' First string not founds so return -1.
    Else
        FuncInstrSandwich = ""
    End If
End Function


'-------------------------------------------------------------------------
' Procedure  : Insert
' Author     : Heriberto Mantilla Santamaría
' Input      : Table, Columns, Values()
' OutPut     : None
' Purpose    : To insert the new record
'-------------------------------------------------------------------------
Public Sub Insert(ByVal Table As String, ByVal Columns As String, ByRef Values() As String)

  Dim xTrue As Boolean
  Dim wSplit As Variant
  Dim NameFile As String
  Dim tPos As Long
  Dim FleeFile As Integer
  Dim iPos As Long
  Dim lData As String
  Dim LIndex As Byte
  Dim jData() As String

On Error GoTo myErr

    If (OpenSTable(Table, IIf(Columns = "*", "", Columns)) = True) Then
        If (Columns = "*") Then
            Insert Table, TableNameReturn, Values
            Exit Sub
        End If
        
        wSplit = Split(Columns, ",")
        ReDim jData(TableFieldCount - 1)

        If (UBound(wSplit) = UBound(Values)) Then

            For iPos = 0 To TableFieldCount - 1

                If (iPos <= UBound(wSplit)) Then

                    For tPos = 0 To TableFieldCount - 1

                        If (UCase$(Trim$(wSplit(iPos))) = UCase$(NTables(2, tPos + 1))) Then
                            xTrue = CompareData(TableValues(tPos + 1, 3), Values(iPos), _
                                TableValues(tPos + 1, 1), TableValues(tPos + 1, 4))

                            If (xTrue = False) Then
                                jData(tPos) = Values(iPos)
                                Exit For
                            Else
                                Err.Raise 520, "Insert", "Cannot insert the fields."
                                Exit Sub
                            End If

                        End If
                        
                    Next tPos

                End If
                
            Next iPos

            lData = ""

            For iPos = 0 To TableFieldCount - 1
                lData = lData & IIf(jData(iPos) = "", vbNullString, jData(iPos)) & vbNullChar
            Next iPos

            If (Trim$(lData) <> "") Then
                lData = lData & vbNewLine
                NameFile = AppDir & "data\" & DatabaseName & "\" & Table & ".udt"
                FleeFile = FreeFile
                LIndex = 1
                Dim zsData As String

                zsData = Decrypt(NameFile)
                lData = lData & zsData
                Open NameFile For Binary Access Write As FleeFile
                    Put FleeFile, LIndex, lData
                Close FleeFile
                EncryptFile NameFile
            End If

        End If
    Else
        Err.Raise 520, "Insert", "Cannot insert the fields."
    End If

    Exit Sub
myErr:
    Err.Raise 520, "Insert", "Cannot insert the fields."

End Sub


'-------------------------------------------------------------------------
' Procedure  : Move
' Author     : Heriberto Mantilla Santamaría
' Input      : NumRecords
' OutPut     : None
' Purpose    : Moves the position of the current record
'-------------------------------------------------------------------------
Public Sub Move(ByVal NumRecords As Integer)

    If (bOpenRecordset = True) And (NumRecords > 0) And (NumRecords <= RecordCount) Then
        mActualPos = NumRecords
    Else
        Err.Raise 512, "Move", "Error moving current record."
    End If

End Sub


'-------------------------------------------------------------------------
' Procedure  : MoveFirst
' Author     : Heriberto Mantilla Santamaría
' Input      : None
' OutPut     : None
' Purpose    : To move to the First record
'-------------------------------------------------------------------------
Public Sub MoveFirst()

    If (bOpenRecordset = True) And (RecordCount > 0) Then
        mActualPos = 0
    Else
        Err.Raise 513, "MoveFirst", "Error moving First record."
    End If

End Sub


'-------------------------------------------------------------------------
' Procedure  : MoveLast
' Author     : Heriberto Mantilla Santamaría
' Input      : None
' OutPut     : None
' Purpose    : To move to the Last record
'-------------------------------------------------------------------------
Public Sub MoveLast()

    If (bOpenRecordset = True) And (RecordCount > 0) Then
        mActualPos = RecordCount
    Else
        Err.Raise 514, "MoveLast", "Error moving Last record."
    End If

End Sub


'-------------------------------------------------------------------------
' Procedure  : MoveNext
' Author     : Heriberto Mantilla Santamaría
' Input      : None
' OutPut     : None
' Purpose    : To move to the Next record
'-------------------------------------------------------------------------
Public Sub MoveNext()

    If (bOpenRecordset = True) And (mActualPos < RecordCount) And (RecordCount > 0) Then
        mActualPos = mActualPos + 1
    Else
        Err.Raise 515, "MoveNext", "Error moving Next record."
    End If

End Sub


'-------------------------------------------------------------------------
' Procedure  : MovePrevious
' Author     : Heriberto Mantilla Santamaría
' Input      : None
' OutPut     : None
' Purpose    : To move to the Previous record
'-------------------------------------------------------------------------
Private Sub MovePrevious()

    If (bOpenRecordset = True) And (mActualPos > 1) And (RecordCount > 0) Then
        mActualPos = mActualPos - 1
    Else
        Err.Raise 516, "MovePrevious", "Error moving Previous record."
    End If

End Sub

Public Property Get Password() As String

    Password = m_sPassword

End Property

Public Property Let Password(ByVal sPassword As String)

    m_sPassword = sPassword

End Property


'-------------------------------------------------------------------------
' Property   : RecordCount
' Author     : Heriberto Mantilla Santamaría
' Input      : None
' OutPut     : The total number of records
' Purpose    : to get the total number of records
'-------------------------------------------------------------------------
Public Property Get RecordCount() As Long

    If (bOpenRecordset = True) Then
        RecordCount = mRecordCount
    Else
        RecordCount = -1
    End If

End Property


'-------------------------------------------------------------------------
' Procedure  : Save
' Author     : Heriberto Mantilla Santamaría
' Input      : DataValue(), TableName, [NewEntry]
' OutPut     : None
' Purpose    : To Save the current fields as a new Record
'-------------------------------------------------------------------------
Public Sub Save(ByRef DataValue() As String, _
                ByVal TableName As String, _
                Optional ByVal NewEntry As Boolean = True)

  Dim vTrue    As Boolean
  Dim iCell    As Long
  Dim bData    As String
  Dim FleeFile As Integer
  Dim NameFile As String
  Dim LIndex    As Byte

    On Error GoTo myErr

    If (bOpenDatabase = False) Or (bOpenRecordset = False) Then
        Err.Raise 521, "Save", "Cannot save the fields."
        Exit Sub
    End If

    bData = ""

    If (UBound(DataValue) > 0) And (UBound(DataValue) <= TableFieldCount) Then

        For iCell = 1 To UBound(DataValue)
            vTrue = CompareData(TableValues(iCell, 3), DataValue(iCell), TableValues(iCell, 1), _
                TableValues(iCell, 4))

            If (vTrue = True) Then
                bData = ""
                Exit For
            Else
                bData = bData & IIf(DataValue(iCell) = "", vbNullString, DataValue(iCell)) & _
                    vbNullChar
            End If

        Next iCell

        If (Trim$(bData) <> "") Then
            bData = bData & vbNewLine
            NameFile = AppDir & "data\" & DatabaseName & "\" & TableName & ".udt"
            FleeFile = FreeFile
            LIndex = 1

            If (NewEntry = True) Then
                Dim zsData As String

                zsData = Decrypt(NameFile)
                bData = bData & zsData
            Else
                Kill NameFile
            End If

            Open NameFile For Binary Access Write As FleeFile
                Put FleeFile, LIndex, bData
            Close FleeFile
            EncryptFile NameFile
        End If

    Else
        Err.Raise 521, "Save", "Cannot save the fields."
    End If

    Exit Sub
myErr:
    Err.Raise 521, "Save", "Cannot save the fields."

End Sub

Public Function SelectFields(ByVal Table As String, Optional ByVal Columns As String = "*", _
                             Optional ByVal Where As String = "") As Boolean

  Dim xTrue  As Boolean
  Dim jPos As Integer
  Dim wSplit As Variant
  Dim Where_T As String
  Dim lmPos As Long
  Dim smPos As Long

On Error GoTo myErr
    xTrue = OpenSTable(Table, IIf(Columns = "*", "", Columns))
    SelectFields = False
    Where = Trim$(Where)

    If (xTrue = True) Then

        FileOp = AppDir & "data\" & DatabaseName & "\" & Table & ".udt"
        FileOp = OpenFile(FileOp, True)

        If (xOpenFile = False) Then
            bOpenRecordset = False
            Exit Function
        End If

        xSplit = Split(FileOp, vbNewLine)
        mRecordCount = UBound(xSplit)

        If (Columns = "*") Then
            SelectFields Table, TableNameReturn
            Exit Function
        Else
            wSplit = Split(Columns, ",")
            ReDim lRecord(mRecordCount - 1, UBound(wSplit))
            
            Init
            For iPos = 1 To TableFieldCount
                AddVar TableValues(iPos, 2), TableValues(iPos, 2)
            Next iPos
            
        End If
        
        If (Where <> "") Then
            Where_T = Where
            ExtractLiteral Where_T
            
            If (isAbort = True) Then
                GoTo myErr
                Exit Function
            End If
            
            Eval Where_T
            If (isAbort = True) Then
                GoTo myErr
                Exit Function
            End If
            
        End If
        
        lmPos = -1
        
        For iPos = 0 To mRecordCount - 1
            ySplit = Split(xSplit(iPos), vbNullChar)

            For tPos = 0 To UBound(ySplit) - 1
                For jPos = 0 To UBound(wSplit)

                    If (tPos <= TableFieldCount - 1) Then
                        If (UCase$(Trim$(wSplit(jPos))) = UCase$(NTables(2, tPos + 1))) Then
                            If (Where = "") Then
                                lRecord(iPos, jPos) = CStr(ySplit(tPos))
                                SelectFields = True
                            Else
                                Dim lK As Long, Where_L As String
                                
                                If (tPos = 0) Then
                                    Where_L = Where
                                    For lK = 0 To (TableFieldCount - 1)
                                        If (IsNumeric(CStr(ySplit(lK))) = True) Then
                                            Where_T = FindWhere(Where_L, NTables(2, lK + 1), CStr(ySplit(lK)))
                                        Else
                                            Where_T = FindWhere(Where_L, NTables(2, lK + 1), "'" & CStr(ySplit(lK)) & "'")
                                        End If
                                    
                                        Where_L = Where_T
                                    Next lK
                                
                                    Where_L = Replace$(Where_L, "'", """")
                                    If (EvalSQL(Where_L) = True) Then
                                        SelectFields = True
                                        If (tPos = 0) Then
                                            lmPos = lmPos + 1
                                        End If
                                        
                                        For smPos = 0 To UBound(wSplit)
                                            lRecord(lmPos, smPos) = CStr(ySplit(smPos))
                                        Next smPos
                                    
                                    End If
                                
                                End If
                                
                            End If
                                
                        End If

                    End If
                        
                Next jPos

            Next tPos
                
        Next iPos
        
        If (Where <> "") Then
            mRecordCount = lmPos + 1
        End If
        
    End If

    bOpenRecordset = True
    Exit Function
myErr:
    SelectFields = False
    bOpenRecordset = False
    Err.Raise 522, "SelectFields", "Error Selecting the fields."

End Function


'-------------------------------------------------------------------------
' Property   : TableFieldCount
' Author     : Heriberto Mantilla Santamaría
' Input      : None
' OutPut     : The total number of datatypes in a table
' Purpose    : to get the total number of datatypes in a table
'-------------------------------------------------------------------------
Public Property Get TableFieldCount() As Long

    If (bOpenTable = True) Then
        TableFieldCount = mTFieldCount
    Else
        TableFieldCount = -1
    End If

End Property

Public Property Get TableFileName() As String

    TableFileName = m_TableName

End Property

Private Function TableInDatabase(ByVal TableName As String) As Boolean

  Dim vCell As Integer

    On Error GoTo myErr
    TableInDatabase = False

    For vCell = 0 To UBound(TableNames) - 1

        If (TableNames(vCell) = TableName) Then
            TableInDatabase = True
            Exit For
        End If

    Next vCell
    Exit Function
myErr:
    TableInDatabase = False

End Function

Public Function TableValues(ByVal Row As Integer, ByVal Col As Integer) As String

    On Error GoTo myErr

    If (bOpenTable = True) Then
        TableValues = NTables(Col, Row)
    Else
        TableValues = ""
        Err.Raise 523, "TableValues", "The table is close."
    End If

    Exit Function
myErr:
    TableValues = ""
    Err.Raise 523, "TableValues", "The table is close."

End Function


Private Function TableNameReturn() As String

  Dim i As Long
  Dim sTable As String
  Dim lCount As Long

    TableNameReturn = ""
    sTable = ""
    lCount = TableFieldCount
    For i = 1 To lCount
        If (i < lCount) Then
            sTable = sTable & TableValues(i, 2) & ","
        Else
            sTable = sTable & TableValues(i, 2)
        End If
        
    Next i
    
    TableNameReturn = sTable
    
End Function


Public Function TotalTables() As Long
    
On Error GoTo TotalTablesErr

    TotalTables = UBound(TableNames)
    Exit Function

TotalTablesErr:

    TotalTables = -1
    
End Function


'-------------------------------------------------------------------------
' Procedure  : Update
' Author     : Heriberto Mantilla Santamaría
' Input      : Table, Columns, Values(), Where
' OutPut     : None
' Purpose    : To update the current fields
'-------------------------------------------------------------------------
Public Sub Update(ByVal Table As String, _
                  ByVal Columns As String, _
                  ByRef Values() As String, _
                  ByVal Where As String)
    
    If (OpenSTable(Table, Columns) = True) Then
        
    Else
        Err.Raise 520, "Insert", "Cannot insert the fields."
    End If
        
End Sub

Public Property Let Username(ByVal sUserName As String)

    m_sUserName = sUserName

End Property

Public Property Get Username() As String

    Username = m_sUserName

End Property


'-------------------------------------------------------------------------
' Function   : OpenDatabase
' Author     : Heriberto Mantilla Santamaría
' Input      : Source, [Username], [Password]
' OutPut     : Boolean True/False
' Purpose    : Open specific database
'-------------------------------------------------------------------------
Public Function OpenDatabase(ByVal Source As String, Optional ByVal Username As String = "", _
                             Optional ByVal Password As String = "") As Boolean
  Dim myHMSDatabase As Boolean

    On Error GoTo myErr:
    ' Take database parameters.
    FileOp = OpenFile(Source, True)

    If (xOpenFile = False) Then
        OpenDatabase = False
        bOpenDatabase = False
        Err.Raise 517, "OpenDatabase", "Cannot open the database."
        Exit Function
    End If

    xSplit = Split(FileOp, vbCrLf)
    ReDim TableNames(0)
    myHMSDatabase = False

    For iPos = 0 To UBound(xSplit)

        If (Trim$(xSplit(iPos)) = "-- myHMS Database") Then
            myHMSDatabase = True
        End If

        TempDt = Trim$(RemoveComment(xSplit(iPos), "-- "))

        If (TempDt <> "") Then
            FileOp = FindStructure(TempDt)

            If (FileOp = "CREATE DATABASE") Then
                DatabaseName = Mid$(TempDt, Len(FileOp) + 3, Len(TempDt) - (Len(FileOp) + 4))
            ElseIf (FileOp = "CREATE USER") Then
                m_sUserName = Mid$(TempDt, Len(FileOp) + 3, Len(TempDt) - (Len(FileOp) + 4))
            ElseIf (FileOp = "CREATE PASSWORD") Then
                m_sPassword = Mid$(TempDt, Len(FileOp) + 3, Len(TempDt) - (Len(FileOp) + 4))
            ElseIf (FileOp = "TABLE NAME") Then
                tPos = InStr(1, RTrim$(TempDt), ";", vbBinaryCompare) - 1
                TableNames(UBound(TableNames)) = Mid$(RTrim$(TempDt), Len(FileOp) + 3, tPos - _
                    (Len(FileOp) + 3))
                ReDim Preserve TableNames(UBound(TableNames) + 1)
            End If

        End If
        
    Next iPos

    If (myHMSDatabase = False) Then
        GoTo myErr
    'ElseIf (m_sUserName <> "") And (m_sUserName <> Username) Then
    '    GoTo myErr
    'ElseIf (m_sPassword <> "") And (m_sPassword <> Password) Then
    '    GoTo myErr
    Else
        OpenDatabase = True
        bOpenDatabase = True
        Exit Function
    End If

myErr:
    OpenDatabase = False
    bOpenDatabase = False
    Err.Raise 517, "OpenDatabase", "Cannot open the database."

End Function


'-------------------------------------------------------------------------
' Function   : OpenRecordSet
' Author     : Heriberto Mantilla Santamaría
' Input      : TableName
' OutPut     : Boolean True/False
' Purpose    : Open fields of table
'-------------------------------------------------------------------------
Public Function OpenRecordSet(ByVal TableName As String) As Boolean
    ' Open fields of table.
    On Error GoTo myErr

    If (bOpenDatabase = False) Then
        OpenRecordSet = False
        bOpenRecordset = False
        Err.Raise 518, "OpenRecordSet", "Cannot open Recordset."
        Exit Function
    End If

    OpenRecordSet = False
    bOpenRecordset = False
    FileOp = AppDir & "data\" & DatabaseName & "\" & TableName & ".udt"
    FileOp = OpenFile(FileOp, True)

    If (xOpenFile = False) Or (TableInDatabase(TableName) = False) Then
        OpenRecordSet = False
        bOpenRecordset = False
        Err.Raise 518, "OpenRecordSet", "Cannot open Recordset."
        Exit Function
    End If

    xSplit = Split(FileOp, vbNewLine)
    mRecordCount = UBound(xSplit)

    If (mRecordCount > 0) Then
        ReDim lRecord(mRecordCount - 1, TableFieldCount)

        For iPos = 0 To mRecordCount - 1
            ySplit = Split(xSplit(iPos), vbNullChar)

            For tPos = 0 To UBound(ySplit) - 1
                lRecord(iPos, tPos) = CStr(ySplit(tPos))
                OpenRecordSet = True
                bOpenRecordset = True
            Next tPos

        Next iPos
        m_TableName = TableName
    Else
        OpenRecordSet = True
        bOpenRecordset = True
    End If

    Exit Function
myErr:
    OpenRecordSet = False
    bOpenRecordset = False
    Err.Raise 518, "OpenRecordSet", "Cannot open Recordset."

End Function


'-------------------------------------------------------------------------
' Function   : OpenSTable
' Author     : Heriberto Mantilla Santamaría
' Input      : TableName, [SpecificColumn]
' OutPut     : Boolean True/False
' Purpose    : Open any table in database with specific column if
'              is necessary
'-------------------------------------------------------------------------
Private Function OpenSTable(ByVal TableName As String, Optional ByVal SpecificColumn As String _
                            = "") As Boolean
  Dim lPos As Long

    On Error GoTo myErr

    If (bOpenDatabase = False) Then
        OpenSTable = False
        bOpenTable = False
        Err.Raise 519, "OpenTable", "Cannot open table."
        Exit Function
    End If

    FileOp = AppDir & "data\" & DatabaseName & "\" & TableName & ".dtb"
    FileOp = OpenFile(FileOp, True)
    
    ' Take table parameters.
    If (xOpenFile = False) Or (TableInDatabase(TableName) = False) Then
        OpenSTable = False
        bOpenTable = False
        Err.Raise 519, "OpenTable", "Cannot open table."
        Exit Function
    End If

    xSplit = Split(FileOp, vbCrLf)
    ReDim NTables(6, 1)
    lPos = 1

    For iPos = 0 To UBound(xSplit)
        TempDt = Trim$(RemoveComment(xSplit(iPos), "-- "))

        If (TempDt <> "") Then
            FileOp = FindStructure(TempDt)

            If (FileOp = "TABLE") Then

                For tPos = iPos + 1 To UBound(xSplit) - 1

                    If (xSplit(tPos) = ");") Then
                        Exit For
                    Else
                        ReDim Preserve NTables(6, lPos)
                        ySplit = Split(xSplit(tPos), vbNullChar)
                        FileOp = Trim$(CStr(Replace$(ySplit(0), "`", "")))
                        NTables(1, lPos) = 0
                        NTables(2, lPos) = FileOp
                        NTables(3, lPos) = CStr(ySplit(1))
                        NTables(4, lPos) = ""

                        If (ySplit(2) = "NOT NULL") Then
                            NTables(4, lPos) = "N"
                        ElseIf (Mid$(ySplit(2), 1, 7) = "DEFAULT") Then
                            FileOp = Mid$(ySplit(2), Len("DEFAULT '") + 1, Len(ySplit(2)))
                            NTables(5, lPos) = CStr(Mid$(FileOp, 1, Len(FileOp) - 1))
                        End If

                        If (Mid$(ySplit(3), 1, 7) = "DEFAULT") Then
                            FileOp = Mid$(ySplit(3), Len("DEFAULT '") + 1, Len(ySplit(3)))
                            NTables(5, lPos) = CStr(Mid$(FileOp, 1, Len(FileOp) - 1))
                        ElseIf (Mid$(ySplit(3), 1, 7) = "COMMENT") Then
                            FileOp = Mid$(ySplit(3), Len("COMMENT '") + 1, Len(ySplit(3)))
                            NTables(6, lPos) = CStr(Mid$(FileOp, 1, Len(FileOp) - 1))
                        End If

                        If (UBound(ySplit) = 4) Then
                            If (Mid$(ySplit(4), 1, 7) = "COMMENT") Then
                                FileOp = Mid$(ySplit(4), Len("COMMENT '") + 1, Len(ySplit(4)))
                                NTables(6, lPos) = CStr(Mid$(FileOp, 1, Len(FileOp) - 1))
                            ElseIf (Mid$(ySplit(4), 1, 2) = "PK") Then
                                NTables(1, lPos) = 1
                            End If

                        End If

                        If (UBound(ySplit) = 5) Then
                            NTables(1, lPos) = 1
                        End If

                        lPos = lPos + 1
                    End If

                Next tPos
                Exit For
            End If

        End If
    Next iPos

    OpenSTable = True
    bOpenTable = True
    mTFieldCount = lPos - 1

    If (SpecificColumn <> "") Then
        Dim wSplit As Variant
        Dim cPos   As Integer
        Dim lFind  As Boolean

        wSplit = Split(SpecificColumn, ",")

        For cPos = 0 To UBound(wSplit)
            lFind = False

            For iPos = 1 To mTFieldCount

                If (UCase$(Trim$(wSplit(cPos))) = UCase$(NTables(2, iPos))) Then
                    lFind = True
                    Exit For
                End If

            Next iPos

            If (lFind = False) Then
                OpenSTable = False
                bOpenTable = False
                Err.Raise 519, "OpenTable", "Cannot open table."
                Exit For
            End If

        Next cPos
    End If

    Exit Function
myErr:
    Err.Raise 519, "OpenTable", "Cannot open table."

End Function


'-------------------------------------------------------------------------
' Function   : OpenTable
' Author     : Heriberto Mantilla Santamaría
' Input      : TableName
'OutPut     : Boolean True/False
' Purpose    : Open any table in database
'-------------------------------------------------------------------------
Public Function OpenTable(ByVal TableName As String) As Boolean
    
    ' Take database parameters.
    OpenTable = OpenSTable(TableName)

End Function
