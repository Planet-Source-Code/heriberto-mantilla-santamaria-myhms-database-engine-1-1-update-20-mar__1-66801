VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About myHMS Database Engine"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin myHMS.FrameMS FrameMS1 
      Height          =   345
      Index           =   0
      Left            =   1065
      Top             =   120
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   609
      BackColor2      =   15329769
      BackColor3      =   16777215
      BackColor4      =   8481373
      BorderColor     =   6582129
      Caption         =   "myHMS Database Engine 1.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      ShadowColor     =   6582129
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3825
      TabIndex        =   0
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) HACKPRO TM 2006"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1935
      Width           =   2730
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0802
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1140
      Index           =   0
      Left            =   1230
      TabIndex        =   1
      Top             =   570
      Width           =   3480
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   105
      Picture         =   "frmAbout.frx":08B4
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdClose_Click()
 Unload frmAbout
 Set frmAbout = Nothing
End Sub
