Attribute VB_Name = "mdlGDI"
Option Explicit

'*******************************************************************
' some font effects, experiment with this if u want some strange font
'*******************************************************************
Public Const DT_BOTTOM As Long = &H8
Public Const DT_CALCRECT As Long = &H400
Public Const DT_CENTER As Long = &H1
Public Const DT_EXPANDTABS As Long = &H40
Public Const DT_HIDEPREFIX As Long = &H100000
Public Const DT_LEFT As Long = &H0
Public Const DT_NOCLIP As Long = &H100
Public Const DT_NOPREFIX As Long = &H800
Public Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
Public Const DT_TOP As Long = &H0
Public Const DT_VCENTER As Long = &H4
Public Const DT_RIGHT As Long = &H2

'*******************************************************************
' change charset for CrFont function, use whatever you wish, no more
' stupid ansi charset ;), fill datagrid with whatever letters u want
'*******************************************************************
Public Const ANSI_CHARSET As Long = 0
Public Const EASTEUROPE_CHARSET As Long = 238
Public Const CHINESEBIG5_CHARSET As Long = 136
Public Const BALTIC_CHARSET As Long = 186
Public Const DEFAULT_CHARSET As Long = 1
Public Const MAC_CHARSET As Long = 77
Public Const OEM_CHARSET As Long = 255
Public Const RUSSIAN_CHARSET As Long = 204
Public Const SYMBOL_CHARSET As Long = 2
Public Const TURKISH_CHARSET As Long = 162
Public Const OUT_DEFAULT_PRECIS As Long = 0
Public Const OUT_STROKE_PRECIS As Long = 3
Public Const PROOF_QUALITY As Long = 2
Public Const DEFAULT_PITCH As Long = 0
Public Const CLIP_DEFAULT_PRECIS As Long = 0
Public Const FF_DONTCARE As Long = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_HEAVY = 900
Public Const FW_BLACK = FW_HEAVY
Public Const FW_DEMIBOLD = FW_SEMIBOLD
Public Const FW_REGULAR = FW_NORMAL
Public Const FW_ULTRABOLD = FW_EXTRABOLD
Public Const FW_ULTRALIGHT = FW_EXTRALIGHT
Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90

Public Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(1 To 32) As Byte
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type POINTAPI
   X As Long
   Y As Long
End Type


Public Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long



Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

Public Function Pt2Rt(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As RECT
   Dim tmp As RECT
   With tmp
      .Left = X1
      .Top = Y1
      .Right = X2
      .Bottom = Y2
   End With
   Pt2Rt = tmp
End Function

Function FromStdFont(Optional sFont As String = "Courier New", _
                     Optional nSize As Integer = 8, _
                     Optional nDegrees As Long = 0, _
                     Optional bBold As Boolean = False, _
                     Optional bItalic As Boolean = False, _
                     Optional bUnderline As Boolean = False, _
                     Optional lCharset As Long = EASTEUROPE_CHARSET) As Long
                     
   Dim tDC&: tDC = GetDC(0)
   Dim fW%: fW = FW_NORMAL
   Dim plf As LOGFONT
   Dim i&
   
   If bBold Then fW = FW_BOLD
   
   FromStdFont = CreateFont(-MulDiv(nSize, GetDeviceCaps(tDC, LOGPIXELSY), 72), _
       0, nDegrees * 10, 0, fW, bItalic, bUnderline, False, _
      lCharset, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, _
      PROOF_QUALITY, DEFAULT_PITCH, sFont)
   
   ReleaseDC 0, tDC
End Function

