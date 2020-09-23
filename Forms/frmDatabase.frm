VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "myHMS Database"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   8565
   Icon            =   "frmDatabase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin vbAcceleratorSGrid6.vbalGrid nspGrid1 
      Height          =   1695
      Left            =   210
      TabIndex        =   24
      Top             =   7080
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   2990
      GridLines       =   -1  'True
      GridLineMode    =   1
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      SelectionAlphaBlend=   -1  'True
   End
   Begin vbAcceleratorSGrid6.vbalGrid nspGrid 
      Height          =   4995
      Left            =   4365
      TabIndex        =   14
      Top             =   405
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8811
      GridLines       =   -1  'True
      NoVerticalGridLines=   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      HighlightSelectedIcons=   0   'False
      SelectionAlphaBlend=   -1  'True
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Fie&lds"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2670
      TabIndex        =   20
      Top             =   6195
      Width           =   1230
   End
   Begin VB.CommandButton cmdInsertTo 
      Caption         =   "Insert In&to"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1455
      TabIndex        =   21
      Top             =   6195
      Width           =   1230
   End
   Begin VB.Frame framOpc5 
      Caption         =   "Fields in ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   165
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   8265
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert Data »"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5220
      TabIndex        =   19
      Top             =   6195
      Width           =   1410
   End
   Begin MSComctlLib.ImageList imgLstIcons 
      Left            =   4050
      Top             =   6195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":0802
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":1014
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtEdit 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4980
      TabIndex        =   15
      Top             =   4125
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton cmdSaveFields 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7155
      TabIndex        =   16
      Top             =   5490
      Width           =   1185
   End
   Begin VB.Frame framOpc4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   120
      TabIndex        =   17
      Top             =   5925
      Width           =   8265
   End
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "Save &Database"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6765
      TabIndex        =   18
      Top             =   6195
      Width           =   1575
   End
   Begin VB.PictureBox picGround 
      BorderStyle     =   0  'None
      Height          =   1230
      Left            =   225
      ScaleHeight     =   1230
      ScaleWidth      =   3870
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   375
      Width           =   3870
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   825
         Width           =   2175
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   255
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   255
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   855
         Width           =   855
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   510
         Width           =   915
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   135
         Width           =   1380
      End
   End
   Begin VB.Frame framOpc3 
      Caption         =   "Tables Fields Structure"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4380
      TabIndex        =   13
      Top             =   120
      Width           =   4020
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   255
      TabIndex        =   12
      Top             =   5460
      Width           =   1185
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   5460
      Width           =   1185
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2865
      TabIndex        =   10
      Top             =   5460
      Width           =   1185
   End
   Begin VB.ListBox lstTables 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   195
      TabIndex        =   9
      Top             =   1965
      Width           =   3900
   End
   Begin VB.Frame framOpc2 
      Caption         =   "Tables"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   1695
      Width           =   4050
   End
   Begin VB.Frame framOpc1 
      Caption         =   "Database Properties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4050
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   22
      Top             =   6195
      Width           =   1230
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Database"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Database"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Database"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuItems 
      Caption         =   "&Items"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Del"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp1 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

 Private Engine    As myHMSEngine
 
 Private NameTable As String
 Private TakeDir   As String
 Private isNew     As Boolean
 Private isNewT    As Boolean

 Private xData     As String
 Private iCell     As Long
 Private jCell     As Long
 
 Private Const CommentSchema  As String = CommentToken & vbCrLf & "-- Create schema "
 
Private Sub cmdDelete_Click()
On Error GoTo myErr:
  ' Del table.
  If (lstTables.ListIndex >= 0) And (lstTables.ListCount > 0) Then
    Kill lstTables.List(lstTables.ListIndex) & ".udt"
    Kill lstTables.List(lstTables.ListIndex) & ".dtb"
    lstTables.RemoveItem lstTables.ListIndex
    isNewT = True
    cmdSaveAll_Click
  End If
  Exit Sub
myErr:
End Sub

Private Sub cmdEdit_Click()
  ' Edit table (Not Finish Sub).
  If (lstTables.ListIndex >= 0) And (lstTables.ListCount > 0) Then
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    nspGrid.Editable = True
    cmdSaveFields.Enabled = True
    lstTables.Enabled = False
    NameTable = lstTables.List(lstTables.ListIndex)
    If (nspGrid.Rows = 0) Then
      nspGrid.AddRow
      nspGrid.CellIcon(nspGrid.Rows, 1) = 0
      nspGrid.CellItemData(nspGrid.Rows, 1) = 0
    End If
  End If
End Sub

Private Sub cmdInsert_Click()
  If (cmdInsert.Caption = "&Insert Data »") Then
   cmdInsert.Caption = "&Insert Data «"
   Me.Height = 9690
   framOpc5.Caption = "Total Fields: " & HMSEngine.RecordCount & " in " & lstTables.List(lstTables.ListIndex)
   framOpc5.Visible = True
  Else
   cmdInsert.Caption = "&Insert Data »"
   Me.Height = 7515
   framOpc5.Caption = "Fields in ..."
   framOpc5.Visible = False
  End If
  Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub cmdInsertTo_Click()
  Dim Values(1) As String

  ' Insert New data
  Values(1) = InputBox("Enter the code", "myHMS Database", "")
  Values(0) = InputBox("Enter the username", "myHMS Database", "")
  HMSEngine.Insert "User", "Username, Code", Values
End Sub

Private Sub cmdNew_Click()
  Dim xTable As String

  ' Create new table.
  xTable = Trim$(InputBox("Write Table Name:", "New table", ""))
  If (xTable <> "") And (FindInList(xTable) = False) Then
    lstTables.AddItem xTable
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    nspGrid.Editable = True
    cmdSaveFields.Enabled = True
    lstTables.Enabled = False
    NameTable = xTable
    xData = ""
    xData = BFinalToken & vbCrLf & CommentToken & "TABLE " & NameTable & vbCrLf & BFinalToken
    CreateFile TakeDir & NameTable & ".dtb", xData
    CreateFile TakeDir & NameTable & ".udt", ""
    nspGrid.Clear
    If (nspGrid.Rows = 0) Then
      nspGrid.AddRow
      nspGrid.CellIcon(nspGrid.Rows, 1) = 0
      nspGrid.CellItemData(nspGrid.Rows, 1) = 0
    End If
    isNewT = True
    lstTables.ListIndex = lstTables.NewIndex
  Else
    NameTable = ""
  End If
  isNewT = True
  cmdSaveAll_Click
End Sub

Private Sub cmdSaveAll_Click()
  Dim TempVal As String

On Error Resume Next
  ' Save the database.
  If (Trim$(txtName.Text) = "") Then
    MsgBox "Error in Database name.", vbCritical + vbOKOnly, "myHMS Database"
    txtName.SetFocus
    Exit Sub
  End If
  xData = BFinalToken & vbCrLf & CommentToken & "myHMS Database" & vbCrLf & BFinalToken
  xData = xData & vbCrLf & vbCrLf & CommentSchema & Trim$(txtName.Text) & _
          vbCrLf & CommentToken & vbCrLf & vbCrLf & "CREATE DATABASE `" & Trim$(txtName.Text) & "`;"
  xData = xData & vbCrLf & vbCrLf & CommentToken & vbCrLf & CommentToken & "CREATE USER & PASSWORD" & _
          vbCrLf & CommentToken & vbCrLf & vbCrLf & "CREATE USER `" & Trim$(txtUser.Text) & "`;" & vbCrLf & _
          "CREATE PASSWORD `" & Trim$(txtPassword.Text) & "`;"
  ' Create Empty dir in data folder.
  TakeDir = AppDir & "data/" & Trim$(txtName.Text) & "/"
  If (FileExits(TakeDir) = False) Then
    MkDir TakeDir
  End If
  ' Create tables in principal file.
  xData = xData & vbCrLf & vbCrLf & CommentToken & vbCrLf & CommentToken & "CREATE TABLES" & vbCrLf & CommentToken & vbCrLf & vbCrLf
  For iCell = 0 To lstTables.ListCount - 1
   If (iCell < lstTables.ListCount - 1) Then
     TempVal = vbCrLf
   Else
     TempVal = ""
   End If
   xData = xData & "TABLE NAME `" & lstTables.List(iCell) & "`;" & TempVal
  Next iCell
  ' Kill database structure.
  Kill TakeDir & Trim$(txtName.Text) & ".dms"
  ' Create principal file.
  CreateFile TakeDir & Trim$(txtName.Text) & ".dms", xData
  If (isNew = True) Then
    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    nspGrid.Editable = False
    cmdSaveFields.Enabled = False
    lstTables.Enabled = True
    cmdSaveAll.Enabled = False
  End If
  isNew = False
On Error GoTo 0
End Sub

Private Sub cmdSaveFields_Click()
  Dim TempVal As String, PK As String
  Dim kData   As String
  Dim NNull   As String
  
 On Error Resume Next
  cmdNew.Enabled = True
  cmdEdit.Enabled = True
  cmdDelete.Enabled = True
  nspGrid.Editable = False
  cmdSaveFields.Enabled = False
  ' Save in the array.
  xData = ""
  xData = BFinalToken & vbCrLf & CommentToken & "TABLE " & NameTable & vbCrLf & BFinalToken & vbCrLf & vbCrLf
  xData = xData & "TABLE `" & NameTable & "`(" & vbCrLf
  For iCell = 1 To nspGrid.Rows
    If (iCell <= nspGrid.Rows) Then
      kData = ","
    Else
      kData = ""
    End If
    If (InDataType(nspGrid.CellText(iCell, 3)) = True) And (Trim$(nspGrid.CellText(iCell, 2)) <> "") Then
      NNull = IIf(nspGrid.CellText(iCell, 4) = "", vbNullChar, vbNullChar & "NOT NULL" & vbNullChar)
      PK = IIf(nspGrid.CellItemData(iCell, 1) = 0, vbNullChar, vbNullChar & "PK")
      TempVal = "`" & nspGrid.CellText(iCell, 2) & "`" & vbNullChar & nspGrid.CellText(iCell, 3) & _
                NNull & "DEFAULT '" & nspGrid.CellText(iCell, 5) & "'" & vbNullChar & "COMMENT '" & nspGrid.CellText(iCell, 6) & "'" & PK & kData
      xData = xData & TempVal & vbCrLf
    End If
  Next iCell
  xData = xData & ");"
  ' Kill file table first.
  Kill TakeDir & NameTable & ".dtb"
  CreateFile TakeDir & NameTable & ".dtb", xData
  lstTables.Enabled = True
On Error GoTo 0
End Sub

Private Sub cmdSelect_Click()
  Dim iPos As Long, tPos As Long
  
  ' Demo Selection
  If (HMSEngine.SelectFields("User", "Code, Username") = True) Then
    If (HMSEngine.RecordCount > 0) Then
      HMSEngine.MoveFirst
    End If
    nspGrid1.Clear True
    nspGrid1.AddColumn , "Code"
    nspGrid1.AddColumn , "Username"
    For iPos = 0 To HMSEngine.RecordCount - 1
      nspGrid1.AddRow
      For tPos = 1 To nspGrid1.Columns
        nspGrid1.CellText(iPos + 1, tPos) = HMSEngine.Fields(tPos - 1)
      Next tPos
      HMSEngine.MoveNext
    Next iPos
    Me.Height = 9690
    framOpc5.Caption = "Total Fields: " & HMSEngine.RecordCount & " in User"
    framOpc5.Visible = True
    Me.Top = (Screen.Height - Me.Height) / 2
  End If
End Sub

Private Sub Form_Load()
  ReDim TableNames(0)
  Set HMSEngine = New myHMSEngine
  isNew = True
  lstTables.Clear
  Me.Height = 7515
  cmdInsert.Enabled = False
  With nspGrid
    .Redraw = False
    .ImageList = imgLstIcons
    Call .AddColumn("PKey", "PK", , , , , , , False)
    Call .AddColumn("ColName", "Column Name")
    Call .AddColumn("DType", "Data type")
    Call .AddColumn("NNull", "Not Null")
    Call .AddColumn("DefVal", "Default Value")
    Call .AddColumn("Comm", "Comment")
    .RowMode = True
    .ScrollBarStyle = ecgSbrRegular
    .SelectionAlphaBlend = True
    .GridLines = True
    .Editable = False
    .HeaderFlat = False
    .HeaderButtons = False
    .ColumnWidth("PKey") = 30
    .ColumnWidth("ColName") = 100
    .ColumnWidth("DType") = 100
    .ColumnWidth("NNull") = 60
    .ColumnWidth("DefVal") = 100
    .ColumnWidth("Comm") = 270
    .ColumnFixedWidth("PKey") = True
    .HighlightSelectedIcons = False
    .StretchLastColumnToFit = False
    .SelectionOutline = True
    .Clear
    .AddRow 0
    .CellIcon(.Rows, 1) = 0
    .CellItemData(.Rows, 1) = 0
    .Redraw = True
  End With
  With nspGrid1
    .Redraw = False
    .RowMode = True
    .ScrollBarStyle = ecgSbrRegular
    .SelectionAlphaBlend = True
    .GridLines = True
    .Editable = False
    .HeaderFlat = False
    .HeaderButtons = False
    .HighlightSelectedIcons = False
    .StretchLastColumnToFit = False
    .SelectionOutline = False
    .Clear
    .Redraw = True
  End With
End Sub

Private Sub lstTables_Click()
  If (isNewT = False) And (lstTables.ListIndex >= 0) And (lstTables.ListCount > 0) Then
    Dim FileOp As String, iPos As Long
    Dim tPos   As Long

    ' Take database parameters.
    FileOp = lstTables.List(lstTables.ListIndex)
    HMSEngine.CloseTable
    HMSEngine.OpenTable FileOp
    nspGrid.Clear
    nspGrid1.Clear True
    For iPos = 1 To HMSEngine.TableFieldCount
      nspGrid1.AddColumn , HMSEngine.TableValues(iPos, 2), ecgHdrTextALignLeft
      nspGrid.AddRow
      nspGrid.CellText(nspGrid.Rows, 2) = HMSEngine.TableValues(iPos, 2)
      nspGrid.CellText(nspGrid.Rows, 3) = HMSEngine.TableValues(iPos, 3)
      nspGrid.CellText(nspGrid.Rows, 4) = HMSEngine.TableValues(iPos, 4)
      nspGrid.CellText(nspGrid.Rows, 5) = HMSEngine.TableValues(iPos, 5)
      nspGrid.CellText(nspGrid.Rows, 6) = HMSEngine.TableValues(iPos, 6)
      nspGrid.CellIcon(nspGrid.Rows, 1) = CInt(HMSEngine.TableValues(iPos, 1))
      nspGrid.CellItemData(nspGrid.Rows, 1) = CInt(HMSEngine.TableValues(iPos, 1))
    Next iPos
    If (nspGrid.Rows = 0) Then
      nspGrid.AddRow
      nspGrid.CellIcon(nspGrid.Rows, 1) = 0
      nspGrid.CellItemData(nspGrid.Rows, 1) = 0
    End If
    cmdInsert.Enabled = True
    mnuDel.Enabled = False
    mnuEdit.Enabled = False
    mnuSave.Enabled = False
    If (HMSEngine.OpenRecordSet(lstTables.List(lstTables.ListIndex)) = True) Then
      If (HMSEngine.RecordCount > 0) Then
        HMSEngine.MoveFirst
      End If
      For iPos = 0 To HMSEngine.RecordCount - 1
        nspGrid1.AddRow
        For tPos = 1 To nspGrid1.Columns
          nspGrid1.CellText(iPos + 1, tPos) = HMSEngine.Fields(tPos - 1)
        Next tPos
        HMSEngine.MoveNext
      Next iPos
    End If
    framOpc5.Caption = "Total Fields: " & HMSEngine.RecordCount & " in " & lstTables.List(lstTables.ListIndex)
  Else
    cmdInsert.Enabled = False
  End If
  isNewT = False
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub mnuAdd_Click()
  ' Add New value.
  nspGrid1.AddRow
  nspGrid1.Editable = True
  mnuSave.Enabled = True
  mnuDel.Enabled = True
  mnuEdit.Enabled = True
End Sub

Private Sub mnuClose_Click()
  HMSEngine.CloseDatabase
  Me.Caption = "myHMS Database"
  lstTables.Clear
  nspGrid.Clear
  nspGrid1.Clear True
  nspGrid.Editable = False
  txtUser.Enabled = False
  txtPassword.Enabled = False
  txtName.Enabled = False
  cmdSaveAll.Enabled = False
  cmdSelect.Enabled = False
  cmdInsertTo.Enabled = False
  cmdNew.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  mnuClose.Enabled = False
  framOpc5.Caption = "Fields in ..."
  txtName.Text = ""
  txtPassword.Text = ""
  txtUser.Text = ""
End Sub

Private Sub mnuDel_Click()
  nspGrid1.Editable = False
  If (nspGrid1.Rows > 0) And (nspGrid1.SelectedRow > 0) Then
    nspGrid1.RemoveRow nspGrid1.SelectedRow
    mnuSave.Enabled = True
  Else
    mnuDel.Enabled = False
    mnuEdit.Enabled = False
  End If
End Sub

Private Sub mnuEdit_Click()
  If (nspGrid1.Rows > 0) And (nspGrid1.SelectedRow > 0) Then
    nspGrid1.Editable = True
    mnuSave.Enabled = True
  End If
End Sub

Private Sub mnuNew_Click()
  ' New database.
  HMSEngine.DatabaseName = ""
  HMSEngine.Username = ""
  HMSEngine.Password = ""
  txtUser.Enabled = True
  txtPassword.Enabled = True
  txtName.Enabled = True
  txtName.SetFocus
  cmdSaveAll.Enabled = True
  cmdSelect.Enabled = True
  cmdInsertTo.Enabled = True
  mnuClose.Enabled = True
End Sub

Private Sub mnuOpen_Click()
  Dim FileOp As String
  Dim xSplit As Variant
  Dim TempDt As String
  Dim tPos   As Long

  ' Open database.
  FileOp = Trim$(ShowOpen(frmDatabase.hWnd))
  If (FileOp <> "") Then
    ' Take database parameters.
    lstTables.Clear
    If (HMSEngine.OpenDatabase(FileOp) = False) Then
      cmdNew.Enabled = False
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
      nspGrid.Editable = False
      cmdSaveFields.Enabled = False
      lstTables.Enabled = False
      TakeDir = ""
      txtName.Text = ""
      txtUser.Text = ""
      txtPassword.Text = ""
      frmDatabase.Caption = "myHMS Database"
      lstTables.Clear
      cmdSaveAll.Enabled = False
      mnuClose.Enabled = False
      cmdSelect.Enabled = False
      cmdInsertTo.Enabled = False
    Else
      cmdNew.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
      nspGrid.Editable = False
      cmdSaveFields.Enabled = False
      lstTables.Enabled = True
      cmdSaveAll.Enabled = True
      isNew = False
      TakeDir = AppDir & "data/" & HMSEngine.DatabaseName & "/"
      txtName.Text = HMSEngine.DatabaseName
      txtUser.Text = HMSEngine.Username
      txtPassword.Text = HMSEngine.Password
      frmDatabase.Caption = "myHMS Database " & HMSEngine.DatabaseName
      If (UBound(TableNames) > 0) Then
        For iCell = 0 To UBound(TableNames) - 1
          lstTables.AddItem TableNames(iCell)
        Next iCell
      End If
      mnuClose.Enabled = True
      cmdSelect.Enabled = True
      cmdInsertTo.Enabled = True
    End If
  End If
End Sub

Private Sub mnuSave_Click()
  Dim lValue() As String
  Dim kPos     As Long
 
  ' Save the data in the specific table.
  mnuSave.Enabled = False
  nspGrid1.Editable = False
  kPos = HMSEngine.RecordCount + 1
  If (kPos <= 0) Then kPos = 1
  For iCell = kPos To nspGrid1.Rows
    ReDim lValue(nspGrid1.Columns)
    For jCell = 1 To nspGrid1.Columns
      lValue(jCell) = nspGrid1.CellText(iCell, jCell)
    Next jCell
    HMSEngine.Save lValue, lstTables.List(lstTables.ListIndex)
  Next iCell
End Sub

Private Sub nspGrid_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, NewValue As Variant, bStayInEditMode As Boolean)
On Error Resume Next
  If (Trim$(txtEdit.Text) = "") And (txtEdit.Visible = True) Then
  Else
    nspGrid.CellText(nspGrid.EditRow, nspGrid.EditCol) = Trim$(txtEdit.Text)
    If (nspGrid.SelectedRow = nspGrid.Rows) Then
      nspGrid.AddRow
      nspGrid.CellIcon(nspGrid.Rows, 1) = 0
      nspGrid.CellItemData(nspGrid.Rows, 1) = 0
    End If
  End If
  Exit Sub
On Error GoTo 0
End Sub

Private Sub nspGrid_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
  Dim lLeft  As Long, lTop    As Long
  Dim lWidth As Long, lHeight As Long
  Dim sText  As String
    
On Error Resume Next
  If (nspGrid.ColumnKey(lCol) = "NNull") Then
    If (nspGrid.CellText(nspGrid.SelectedRow, 4) = "") Then
      nspGrid.CellText(nspGrid.SelectedRow, 4) = "N"
    Else
      nspGrid.CellText(nspGrid.SelectedRow, 4) = ""
    End If
    bCancel = True
    Exit Sub
  ElseIf (nspGrid.ColumnKey(lCol) = "PKey") Then
    If (nspGrid.CellItemData(nspGrid.SelectedRow, 1) = 0) Then
      nspGrid.CellItemData(nspGrid.SelectedRow, 1) = 1
      nspGrid.CellText(nspGrid.SelectedRow, 4) = "N"
      nspGrid.CellIcon(nspGrid.SelectedRow, 1) = 1
    Else
      nspGrid.CellItemData(nspGrid.SelectedRow, 1) = 0
      nspGrid.CellText(nspGrid.SelectedRow, 4) = ""
      nspGrid.CellIcon(nspGrid.SelectedRow, 1) = 0
    End If
    bCancel = True
    Exit Sub
  End If
  '* Tomo el ancho de la celda.
  Call nspGrid.CellBoundary(lRow, lCol, lLeft, lTop, lWidth, lHeight)
  '* Tomo el texto de la celda.
  If Not (IsMissing(nspGrid.CellText(lRow, lCol)) = True) Then
    sText = nspGrid.CellFormattedText(lRow, lCol)
  Else
    sText = ""
  End If
  txtEdit.Text = Trim$(sText)
  If (nspGrid.CellBackColor(lRow, lCol) = -1) Then
    txtEdit.BackColor = nspGrid.BackColor
  Else
    txtEdit.BackColor = nspGrid.CellBackColor(lRow, lCol)
  End If
  ' Si el tamaño del Grid cambia debo cambiar el del cuadro de texto ó combobox.
  Call txtEdit.Move(lLeft + nspGrid.Left + 1, lTop + nspGrid.Top + Screen.TwipsPerPixelY, lWidth, lHeight)
  txtEdit.Visible = True
  txtEdit.ZOrder
  txtEdit.SetFocus
On Error GoTo 0
End Sub

Private Sub nspGrid_ScrollChange(ByVal eBar As vbAcceleratorSGrid6.ECGScrollBarTypes)
  txtEdit.Visible = False
  Call nspGrid.EndEdit
End Sub

Private Sub nspGrid_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
  txtEdit.Visible = False
  Call nspGrid.EndEdit
End Sub

Private Sub nspGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, bDoDefault As Boolean)
  If (Button = 2) Then
    If (nspGrid1.Rows > 0) Then
      mnuEdit.Enabled = True
      mnuDel.Enabled = True
    Else
      mnuEdit.Enabled = False
      mnuDel.Enabled = False
    End If
    Me.PopupMenu mnuItems
  End If
End Sub

Private Sub nspGrid1_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, NewValue As Variant, bStayInEditMode As Boolean)
On Error Resume Next
  nspGrid1.CellText(nspGrid1.EditRow, nspGrid1.EditCol) = Trim$(txtEdit.Text)
On Error GoTo 0
End Sub

Private Sub nspGrid1_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
  Dim lLeft  As Long, lTop    As Long
  Dim lWidth As Long, lHeight As Long
  Dim sText  As String
    
On Error Resume Next
  '* Tomo el ancho de la celda.
  Call nspGrid1.CellBoundary(lRow, lCol, lLeft, lTop, lWidth, lHeight)
  '* Tomo el texto de la celda.
  If Not (IsMissing(nspGrid1.CellText(lRow, lCol)) = True) Then
    sText = nspGrid1.CellFormattedText(lRow, lCol)
  Else
    sText = ""
  End If
  txtEdit.Text = Trim$(sText)
  If (nspGrid1.CellBackColor(lRow, lCol) = -1) Then
    txtEdit.BackColor = nspGrid1.BackColor
  Else
    txtEdit.BackColor = nspGrid1.CellBackColor(lRow, lCol)
  End If
  ' Si el tamaño del Grid cambia debo cambiar el del cuadro de texto ó combobox.
  Call txtEdit.Move(lLeft + nspGrid1.Left + 1, lTop + nspGrid1.Top + Screen.TwipsPerPixelY, lWidth, lHeight)
  txtEdit.Visible = True
  txtEdit.ZOrder
  txtEdit.SetFocus
On Error GoTo 0
End Sub

Private Sub nspGrid1_ScrollChange(ByVal eBar As vbAcceleratorSGrid6.ECGScrollBarTypes)
  txtEdit.Visible = False
  Call nspGrid1.EndEdit
End Sub

Private Sub nspGrid1_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
  txtEdit.Visible = False
  Call nspGrid1.EndEdit
End Sub

Private Function FindInList(ByVal xTable As String) As Boolean
  ' Find Item in a listbox.
  FindInList = False
  For iCell = 0 To lstTables.ListCount - 1
    If (lstTables.List(iCell) = xTable) Then
      FindInList = True
      Exit For
    End If
  Next iCell
End Function

