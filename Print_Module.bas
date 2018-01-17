Attribute VB_Name = "Print_Module"
Public Declare Function StartDoc Lib "gdi32" Alias _
 "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Public Declare Function EndDoc Lib "gdi32" _
 (ByVal hdc As Long) As Long
Public Declare Function StartPage Lib "gdi32" _
 (ByVal hdc As Long) As Long
Public Declare Function EndPage Lib "gdi32" _
 (ByVal hdc As Long) As Long
'Public Declare Function CreateFontIndirect Lib "gdi32" _
 Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" _
 (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" _
  (ByVal hObject As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias _
"TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal _
y As Long, ByVal lpString As String, ByVal nCount _
As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" _
 (ByVal hdc As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" _
 (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" _
 (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function Rectangle Lib "gdi32" _
 (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
 ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" _
 (ByVal hdc As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" _
 (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function LoadImage Lib "User32.dll" _
 Alias "LoadImageA" (ByVal hInst As Long, ByVal _
 lpsz As String, ByVal un1 As Long, ByVal n1 As Long, _
 ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" _
 (ByVal hdc As Long, _
 ByVal x As Long, ByVal y As Long, _
 ByVal nWidth As Long, ByVal nHeight As Long, _
 ByVal hSrcDC As Long, ByVal xSrc As Long, _
 ByVal ySrc As Long, _
 ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
 ByVal dwRop As Long) As Long
 
 'Public Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
 
Public Const LF_FACESIZE = 32
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const PHYSICALOFFSETX = 112
Public Const PHYSICALOFFSETY = 113
Public Const PHYSICALWIDTH = 110
Public Const PHYSICALHEIGHT = 111
Public Const POINTSPERINCH = 72
Public Const NORMAL = 400
Public Const Bold = 700
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2
Public Const IMAGE_BITMAP As Long = &H0
Public Const LR_LOADFROMFILE As Long = &H10
Public Const LR_CREATEDIBSECTION As Long = &H2000
Public Const STRETCH_ANDSCANS = 1
Public Const STRETCH_ORSCANS = 2
Public Const STRETCH_DELETESCANS = 3
Public Const STRETCH_HALFTONE = 4
Public Type DOCINFO
 cbSize As Long
 lpszDocName As String
 lpszOutput As String
 lpszDatatype As String
 fwType As Long
End Type


Public Type PrinterInfo
 Handle As Long
 PixPerInchX As Long
 PixPerInchY As Long
 OffsetX As Long
 OffsetY As Long
 PageWidthInches As Single
 PageHeightInches As Single
End Type
' *** END OF MODULE CODE ***

