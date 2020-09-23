VERSION 5.00
Begin VB.UserControl FrameMS 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
End
Attribute VB_Name = "FrameMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  Private Type OSVERSIONINFO
     dwOSVersionInfoSize As Long
     dwMajorVersion      As Long
     dwMinorVersion      As Long
     dwBuildNumber       As Long
     dwPlatformId        As Long
     szCSDVersion        As String * 128 '* Maintenance string for PSS usage.
  End Type

  Private Type BITMAPINFOHEADER
     biSize          As Long
     biWidth         As Long
     biHeight        As Long
     biPlanes        As Integer
     biBitCount      As Integer
     biCompression   As Long
     biSizeImage     As Long
     biXPelsPerMeter As Long
     biYPelsPerMeter As Long
     biClrUsed       As Long
     biClrImportant  As Long
  End Type

  Private Type POINTAPI
     X           As Long
     Y           As Long
  End Type
  
  Private Type RECT
     lLeft       As Long
     lTop        As Long
     lRight      As Long
     lBottom     As Long
  End Type
  
   Public Enum GradientDirectionEnum
     [Fill_None] = 0
     [Fill_Horizontal] = 1
     [Fill_HorizontalMiddleOut] = 2
     [Fill_Vertical] = 3
     [Fill_VerticalMiddleOut] = 4
     [Fill_DownwardDiagonal] = 5
     [Fill_UpwardDiagonal] = 6
  End Enum
  
  '* Declares for Unicode support.
  Private Const VER_PLATFORM_WIN32_NT = 2
  
  Private Const COLOR_GRAYTEXT   As Long = 17
  Private Const defBackColor1    As Long = vbButtonFace
  Private Const defBackColor2    As Long = &HF6D9A6
  Private Const defBackColor3    As Long = &H60C6FF
  Private Const defBorderColor   As Long = &HBC9A80
  Private Const defShadowColor   As Long = &HC7B9AB
  Private Const defBackColor4    As Long = &H26A5FF
  Private Const DIB_RGB_ColS     As Long = 0
  Private Const DT_LEFT          As Long = &H0
  Private Const DT_SINGLELINE    As Long = &H20
  Private Const DT_VCENTER       As Long = &H4
  Private Const DT_WORD_ELLIPSIS As Long = &H40000
    
  Private g_Font          As StdFont
  Private m_lCaption      As String
  Private m_lEnabled      As Boolean
  Private m_btnRect       As RECT
  Private m_lBackColor1   As OLE_COLOR
  Private m_lBackColor2   As OLE_COLOR
  Private m_lBackColor3   As OLE_COLOR
  Private m_lBorderColor  As OLE_COLOR
  Private m_lBackColor4   As OLE_COLOR
  Private m_lForeColor    As OLE_COLOR
  Private m_lShadowColor  As OLE_COLOR
  Private mWindowsNT      As Boolean
  
  Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
  Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
  Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
  Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
  Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
  Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
  Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
  
  ' for Carles P.V DIB solutions
  Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
  Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
  Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
  Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
  Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
  Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long

Private Sub APILine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lColor As Long)
   Dim PT As POINTAPI, hPen As Long, hPenOld As Long
 
   '* Use the API LineTo for Fast Drawing.
   hPen = CreatePen(0, 1, lColor)
   hPenOld = SelectObject(UserControl.hDC, hPen)
   Call MoveToEx(UserControl.hDC, X1, Y1, PT)
   Call LineTo(UserControl.hDC, X2, Y2)
   Call SelectObject(hDC, hPenOld)
   Call DeleteObject(hPen)
End Sub

Private Function ConvertSystemColor(ByVal theColor As Long) As Long
   '* Convert Long to System Color.
   Call OleTranslateColor(theColor, 0, ConvertSystemColor)
End Function

'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : DIBGradient
' Auther    : Carls P.V.
' Input     : Hdc + Parameters
' OutPut    : None
' Purpose   : DIB solution for fast gradients
'------------------------------------------------------------------------------------------------------------------------------------------

Private Sub DIBGradient(ByVal hDC As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal vWidth As Long, _
                         ByVal vHeight As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionEnum)

  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim R1      As Long
  Dim G1      As Long
  Dim B1      As Long
  Dim R2      As Long
  Dim G2      As Long
  Dim B2      As Long
  Dim dR      As Long
  Dim dG      As Long
  Dim dB      As Long
  
  Dim Scan    As Long
  Dim i       As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim j       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  
    '-- A minor check
    If (vWidth < 1 Or vHeight < 1) Then Exit Sub
    
    '-- Decompose Cols'
    R1 = (Col1 And &HFF&)
    G1 = (Col1 And &HFF00&) \ &H100&
    B1 = (Col1 And &HFF0000) \ &H10000
    R2 = (Col2 And &HFF&)
    G2 = (Col2 And &HFF00&) \ &H100&
    B2 = (Col2 And &HFF0000) \ &H10000

    '-- Get Col distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    
    '-- Size gradient-Cols array
    Select Case GradientDirection
        Case [Fill_Horizontal]
            ReDim lGrad(0 To vWidth - 1)
        Case [Fill_Vertical]
            ReDim lGrad(0 To vHeight - 1)
        Case Else
            ReDim lGrad(0 To vWidth + vHeight - 2)
    End Select
    
    '-- Calculate gradient-Cols
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    
    '-- Size DIB array
    ReDim lBits(vWidth * vHeight - 1) As Long
    iEnd = vWidth - 1
    jEnd = vHeight - 1
    Scan = vWidth
    
    '-- Render gradient DIB
    Select Case GradientDirection
        
        Case [Fill_Horizontal]
        
            For j = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next j
        
        Case [Fill_Vertical]
        
            For j = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(j)
                Next i
                iOffset = iOffset + Scan
            Next j
            
        Case [Fill_DownwardDiagonal]
            
            iOffset = jEnd * Scan
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = j
            Next j
            
        Case [Fill_UpwardDiagonal]
            
            iOffset = 0
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = j
            Next j
    End Select
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = vWidth
        .biHeight = vHeight
    End With
    
    '-- Paint it!
    Call StretchDIBits(hDC, X, Y, vWidth, vHeight, 0, 0, vWidth, vHeight, lBits(0), uBIH, DIB_RGB_ColS, vbSrcCopy)

End Sub

Private Sub DrawCaption(ByVal lCaption As String, Optional ByVal lColor As OLE_COLOR = &HF0)
   If (m_lEnabled = False) Then lColor = GetSysColor(COLOR_GRAYTEXT)
   Call SetTextColor(UserControl.hDC, lColor)
   m_btnRect.lBottom = UserControl.ScaleHeight
   m_btnRect.lLeft = 12
   m_btnRect.lTop = 1
   m_btnRect.lRight = UserControl.ScaleWidth
   '*************************************************************************
   '* Draws the text with Unicode support based on OS version.              *
   '* Thanks to Richard Mewett.                                             *
   '*************************************************************************
   If (mWindowsNT = True) Then
     Call DrawTextW(UserControl.hDC, StrPtr(lCaption), Len(lCaption), m_btnRect, DT_VCENTER Or DT_LEFT Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
   Else
     Call DrawTextA(UserControl.hDC, lCaption, Len(lCaption), m_btnRect, DT_VCENTER Or DT_LEFT Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
   End If
End Sub

Private Sub DrawHGradient(ByVal lEndColor As Long, ByVal lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
   Dim dR As Single, dG As Single, dB As Single
   Dim sR As Single, sG As Single, sB As Single
   Dim eR As Single, eG As Single, eB As Single
   Dim lh As Long, lw   As Long, ni   As Long
  
   '* Draw a Horizontal Gradient in the current HDC
   lh = Y2 - Y
   lw = X2 - X
   sR = (lStartcolor And &HFF)
   sG = (lStartcolor \ &H100) And &HFF
   sB = (lStartcolor And &HFF0000) / &H10000
   eR = (lEndColor And &HFF)
   eG = (lEndColor \ &H100) And &HFF
   eB = (lEndColor And &HFF0000) / &H10000
   dR = (sR - eR) / lw
   dG = (sG - eG) / lw
   dB = (sB - eB) / lw

   For ni = 0 To lw
     Call APILine(X + ni, Y, X + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
   Next ni
End Sub


Private Sub DrawVGradient(ByVal lEndColor As Long, ByVal lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
   Dim dR As Single, dG As Single, dB As Single, ni As Long
   Dim sR As Single, sG As Single, sB As Single
   Dim eR As Single, eG As Single, eB As Single
 
   '* Draw a Vertical Gradient in the current hDC.
   sR = (lStartcolor And &HFF)
   sG = (lStartcolor \ &H100) And &HFF
   sB = (lStartcolor And &HFF0000) / &H10000
   eR = (lEndColor And &HFF)
   eG = (lEndColor \ &H100) And &HFF
   eB = (lEndColor And &HFF0000) / &H10000
   dR = (sR - eR) / Y2
   dG = (sG - eG) / Y2
   dB = (sB - eB) / Y2
  For ni = 0 To Y2
    Call APILine(X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
  Next ni
End Sub

Private Sub Refresh()
   Dim iPos As Integer

   '* Refresh the control.
On Error Resume Next
   Height = 345
   BackColor = m_lBackColor1
   Set UserControl.Font = g_Font
   Cls
   Call DIBGradient(hDC, 7, 2, ScaleWidth, ScaleHeight / 2 - 1, m_lBackColor2, ShiftColorOXP(m_lBackColor2, 120), Fill_Vertical)
   For iPos = 3 To 10
     Call DrawHGradient(ShiftColorOXP(m_lBackColor2, iPos * 22), ShiftColorOXP(m_lBackColor2, iPos * 11), 10, iPos, (ScaleWidth / 2) - 20, iPos + 1)
   Next iPos
   
   Call DIBGradient(hDC, 7, 2 + ScaleHeight / 2 - 1, ScaleWidth, ScaleHeight / 2 - 2, ShiftColorOXP(m_lBackColor2, 120), ConvertSystemColor(m_lBackColor2), Fill_Vertical)
   
   Call DIBGradient(hDC, 1, 3 + ScaleHeight / 2 - 1, 6, ScaleHeight / 2 - 4, ShiftColorOXP(m_lBackColor3), m_lBackColor4, Fill_Vertical)
   Call DIBGradient(hDC, 1, 3, 6, ScaleHeight / 2 - 1, m_lBackColor4, ShiftColorOXP(m_lBackColor3), Fill_Vertical)
   
   Call DIBGradient(hDC, 6, ScaleHeight - 4, ScaleWidth - 1, ScaleHeight - 3, ShiftColorOXP(m_lBackColor2), ShiftColorOXP(m_lBackColor2, 50), Fill_Horizontal)
   
   '* Borders
   Call DrawHGradient(m_lBorderColor, m_lBackColor1, 7, 2, ScaleWidth - 1, 1) '* Top horizontal.
   Call DrawHGradient(m_lBorderColor, m_lBackColor1, 7, ScaleHeight - 2, ScaleWidth - 1, ScaleHeight - 1) '* Bottom horizontal.
   
   Call DrawHGradient(m_lBorderColor, ShiftColorOXP(m_lBorderColor, 120), 1, 2, 5, 1) '* Top horizontal.
   Call DrawHGradient(m_lBorderColor, ShiftColorOXP(m_lBorderColor, 120), 1, ScaleHeight - 2, 5, ScaleHeight - 1) '* Bottom horizontal.
   
   Call DrawHGradient(ShiftColorOXP(m_lBackColor2, 20), ShiftColorOXP(m_lBackColor2, 170), ScaleWidth - 65, 4, ScaleWidth, (ScaleHeight / 2) - 6)
   Call DrawHGradient(ShiftColorOXP(m_lBackColor2, 20), ShiftColorOXP(m_lBackColor2, 170), ScaleWidth - 65, (ScaleHeight / 2) + 6, ScaleWidth, ScaleHeight - 3)
   
   Call DrawVGradient(ShiftColorOXP(m_lBackColor2, 50), ShiftColorOXP(m_lBackColor2), 7, 3, 10, ScaleHeight - 6)
   
   Call APILine(1, 2, 1, ScaleHeight - 1, m_lBorderColor) '* Left border.
   Call APILine(6, 2, 6, ScaleHeight - 1, m_lBorderColor) '* Left border.
   
   For iPos = 1 To ScaleWidth
     If (ScaleWidth / 2 <= iPos - 70) Then
       Call APILine(iPos, ScaleHeight - 1, iPos + 1, ScaleHeight - 1, m_lBackColor1)
     Else
       Call APILine(iPos, ScaleHeight - 1, iPos + 1, ScaleHeight - 1, ShiftColorOXP(m_lShadowColor, iPos + 10))
     End If
   Next iPos
   
   Call DrawCaption(m_lCaption, m_lForeColor)
On Error GoTo 0
End Sub

Private Sub SetAccessKeys()
   Dim AmperSandPos As Long

   UserControl.AccessKeys = ""
   If (Len(Caption) > 1) Then
     AmperSandPos = InStr(1, Caption, "&", vbTextCompare)
     If (AmperSandPos < Len(Caption)) And (AmperSandPos > 0) Then
       If (Mid$(Caption, AmperSandPos + 1, 1) <> "&") Then
         UserControl.AccessKeys = LCase$(Mid$(Caption, AmperSandPos + 1, 1))
       Else
         AmperSandPos = InStr(AmperSandPos + 2, Caption, "&", vbTextCompare)
         If (Mid$(Caption, AmperSandPos + 1, 1) <> "&") Then
           UserControl.AccessKeys = LCase$(Mid$(Caption, AmperSandPos + 1, 1))
         End If
       End If
     End If
   End If
End Sub

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
   Dim cRed   As Long, cBlue  As Long
   Dim Delta  As Long, cGreen As Long

   '* Shift a color.
   cBlue = ((theColor \ &H10000) Mod &H100)
   cGreen = ((theColor \ &H100) Mod &H100)
   cRed = (theColor And &HFF)
   Delta = &HFF - Base
   cBlue = Base + cBlue * Delta \ &HFF
   cGreen = Base + cGreen * Delta \ &HFF
   cRed = Base + cRed * Delta \ &HFF
   If (cRed > 255) Then cRed = 255
   If (cGreen > 255) Then cGreen = 255
   If (cBlue > 255) Then cBlue = 255
   ShiftColorOXP = cRed + 256& * cGreen + 65536 * cBlue
End Function

'-------------------------------------------------------------
' PROPERTY'S
'-------------------------------------------------------------

Public Property Get BackColor1() As OLE_COLOR
   BackColor1 = m_lBackColor1
End Property

Public Property Let BackColor1(ByVal New_Color As OLE_COLOR)
   m_lBackColor1 = ConvertSystemColor(New_Color)
   Call PropertyChanged("BackColor1")
   Call Refresh
End Property

Public Property Get BackColor2() As OLE_COLOR
   BackColor2 = m_lBackColor2
End Property

Public Property Let BackColor2(ByVal New_Color As OLE_COLOR)
   m_lBackColor2 = ConvertSystemColor(New_Color)
   Call PropertyChanged("BackColor2")
   Call Refresh
End Property

Public Property Get BackColor3() As OLE_COLOR
   BackColor3 = m_lBackColor3
End Property

Public Property Let BackColor3(ByVal New_Color As OLE_COLOR)
   m_lBackColor3 = ConvertSystemColor(New_Color)
   Call PropertyChanged("BackColor3")
   Call Refresh
End Property

Public Property Get BackColor4() As OLE_COLOR
   BackColor4 = m_lBackColor4
End Property

Public Property Let BackColor4(ByVal New_Color As OLE_COLOR)
   m_lBackColor4 = ConvertSystemColor(New_Color)
   Call PropertyChanged("BackColor4")
   Call Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_lBorderColor
End Property

Public Property Let BorderColor(ByVal New_Color As OLE_COLOR)
   m_lBorderColor = ConvertSystemColor(New_Color)
   Call PropertyChanged("BorderColor")
   Call Refresh
End Property

Public Property Get Caption() As String
   Caption = m_lCaption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_lCaption = New_Caption
   Call PropertyChanged("Caption")
   Call Refresh
End Property

Public Property Get Enabled() As Boolean
   Enabled = m_lEnabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled = New_Enabled
   m_lEnabled = New_Enabled
   Call PropertyChanged("Enabled")
   Call Refresh
End Property

Public Property Get Font() As StdFont
   Set Font = g_Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
On Error Resume Next
   With g_Font
     .Name = New_Font.Name
     .Size = New_Font.Size
     .Bold = New_Font.Bold
     .Italic = New_Font.Italic
     .Underline = New_Font.Underline
     .Strikethrough = New_Font.Strikethrough
   End With
   Call PropertyChanged("Font")
   Call Refresh
On Error GoTo 0
End Property

Public Property Get ForeColor() As OLE_COLOR
   ForeColor = m_lForeColor
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
   m_lForeColor = ConvertSystemColor(NewColor)
   Call PropertyChanged("ForeColor")
   Call Refresh
End Property

Public Property Get hDC() As Long
   hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get ShadowColor() As OLE_COLOR
   ShadowColor = m_lShadowColor
End Property

Public Property Let ShadowColor(ByVal New_Color As OLE_COLOR)
   m_lShadowColor = ConvertSystemColor(New_Color)
   Call PropertyChanged("ShadowColor")
   Call Refresh
End Property

'-------------------------------------------------------------
' INTERNAL USERCONTROL EVENT'S
'-------------------------------------------------------------

Private Sub UserControl_Initialize()
   Dim OS As OSVERSIONINFO

   '* Get the operating system version for text drawing purposes.
   OS.dwOSVersionInfoSize = Len(OS)
   Call GetVersionEx(OS)
   mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
End Sub

Private Sub UserControl_InitProperties()
   m_lBackColor1 = ConvertSystemColor(defBackColor1)
   m_lBackColor2 = ConvertSystemColor(defBackColor2)
   m_lBackColor3 = ConvertSystemColor(defBackColor3)
   m_lBackColor4 = ConvertSystemColor(defBackColor4)
   m_lBorderColor = ConvertSystemColor(defBorderColor)
   m_lCaption = Ambient.DisplayName
   m_lEnabled = True
   m_lForeColor = &H80000012
   m_lShadowColor = defShadowColor
   Set g_Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   BackColor1 = PropBag.ReadProperty("BackColor1", ConvertSystemColor(defBackColor1))
   BackColor2 = PropBag.ReadProperty("BackColor2", ConvertSystemColor(defBackColor2))
   BackColor3 = PropBag.ReadProperty("BackColor3", ConvertSystemColor(defBackColor3))
   BackColor4 = PropBag.ReadProperty("BackColor4", ConvertSystemColor(defBackColor4))
   BorderColor = PropBag.ReadProperty("BorderColor", ConvertSystemColor(defBorderColor))
   m_lCaption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
   Enabled = PropBag.ReadProperty("Enabled", True)
   Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
   ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
   UserControl.MousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
   Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
   Call SetAccessKeys
   ShadowColor = PropBag.ReadProperty("ShadowColor", ConvertSystemColor(defShadowColor))
End Sub

Private Sub UserControl_Resize()
   Call Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("BackColor1", m_lBackColor1, ConvertSystemColor(defBackColor1))
   Call PropBag.WriteProperty("BackColor2", m_lBackColor2, ConvertSystemColor(defBackColor2))
   Call PropBag.WriteProperty("BackColor3", m_lBackColor3, ConvertSystemColor(defBackColor3))
   Call PropBag.WriteProperty("BackColor4", m_lBackColor4, ConvertSystemColor(defBackColor4))
   Call PropBag.WriteProperty("BorderColor", m_lBorderColor, ConvertSystemColor(defBorderColor))
   Call PropBag.WriteProperty("Caption", m_lCaption, Ambient.DisplayName)
   Call PropBag.WriteProperty("Enabled", m_lEnabled, True)
   Call PropBag.WriteProperty("Font", g_Font, Ambient.Font)
   Call PropBag.WriteProperty("ForeColor", m_lForeColor, &H80000012)
   Call PropBag.WriteProperty("MousePointer", MousePointer, vbDefault)
   Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
   Call PropBag.WriteProperty("ShadowColor", m_lShadowColor, ConvertSystemColor(defShadowColor))
End Sub
