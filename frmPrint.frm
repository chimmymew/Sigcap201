VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "การพิมพ์และรายงาน"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2895
   Begin VB.CommandButton Command3 
      Caption         =   "พิมพ์เฉพาะผลการวิเคราะห์"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "พิมพ์กราฟและผลการวิเคราะห์"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "พิมพ์เฉพาะกราฟของสัญญาณ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyPrinter As PrinterInfo
Private UserCancelled As Boolean
Private Myfont As LOGFONT
Private hOldfont As Long
Private hNewFont As Long
Private Declare Function GetDeviceCaps Lib "gdi32" _
(ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" _
Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias _
"RtlMoveMemory" (hpvDest As Any, hpvSource As Any, _
ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" _
(ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" _
(ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" _
(ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" _
(ByVal hMem As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" _
(ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
Private Const PD_PRINTSETUP = &H40
Private Const PD_DISABLEPRINTTOFILE = &H80000
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113
Private Const OPAQUE = 0
Private Const TRANSPARENT = 1
Private Type PRINTDLG_TYPE
lStructSize As Long
hwndOwner As Long
hDevMode As Long
hDevNames As Long
hdc As Long
flags As Long
nFromPage As Integer
nToPage As Integer
nMinPage As Integer
nMaxPage As Integer
nCopies As Integer
hInstance As Long
lCustData As Long
lpfnPrintHook As Long
lpfnSetupHook As Long
lpPrintTemplateName As String
lpSetupTemplateName As String
hPrintTemplate As Long
hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
wDriverOffset As Integer
wDeviceOffset As Integer
wOutputOffset As Integer
wDefault As Integer
extra As String * 200
End Type
Private Type DEVMODE_TYPE
dmDeviceName As String * CCHDEVICENAME
dmSpecVersion As Integer
dmDriverVersion As Integer
dmSize As Integer
dmDriverExtra As Integer
dmFields As Long
dmOrientation As Integer
dmPaperSize As Integer
dmPaperLength As Integer
dmPaperWidth As Integer
dmScale As Integer
dmCopies As Integer
dmDefaultSource As Integer
dmPrintQuality As Integer
dmColor As Integer
dmDuplex As Integer
dmYResolution As Integer
dmTTOption As Integer
dmCollate As Integer
dmFormName As String * CCHFORMNAME
dmUnusedPadding As Integer
dmBitsPerPel As Integer
dmPelsWidth As Long
dmPelsHeight As Long
dmDisplayFlags As Long
dmDisplayFrequency As Long
dmICMMethod As Long
dmICMIntent As Long
dmMediaType As Long
dmDitherType As Long
dmReserved1 As Long
dmReserved2 As Long
dmPanningWidth As Long
dmPanningHeight As Long
End Type

Private Sub SetPrinterOrigin(x As Single, y As Single)
With Printer
.ScaleLeft = .ScaleX(GetDeviceCaps _
(.hdc, PHYSICALOFFSETX), _
vbPixels, .ScaleMode) - x
.ScaleTop = .ScaleY(GetDeviceCaps _
(.hdc, PHYSICALOFFSETY), _
vbPixels, .ScaleMode) - y
.CurrentX = 0
.CurrentY = 0
End With
End Sub

Private Sub GetMyPrinter()
UserCancelled = False
mainMDI.cmd1.DialogTitle = "รายงานและการพิมพ์"
mainMDI.cmd1.PrinterDefault = False
mainMDI.cmd1.CancelError = True
mainMDI.cmd1.Min = 1
mainMDI.cmd1.Max = 2
mainMDI.cmd1.flags = cdlPDReturnDC 'Or cdlPDPrintSetup
On Error GoTo dlgError
mainMDI.cmd1.ShowPrinter
On Error GoTo 0
MyPrinter.Handle = mainMDI.cmd1.hdc
MyPrinter.PixPerInchX = GetDeviceCaps _
 (MyPrinter.Handle, LOGPIXELSX)
MyPrinter.PixPerInchY = GetDeviceCaps _
 (MyPrinter.Handle, LOGPIXELSY)
MyPrinter.OffsetX = GetDeviceCaps _
 (MyPrinter.Handle, PHYSICALOFFSETX)
MyPrinter.OffsetY = GetDeviceCaps _
 (MyPrinter.Handle, PHYSICALOFFSETY)
MyPrinter.PageWidthInches = CSng(GetDeviceCaps _
 (MyPrinter.Handle, PHYSICALWIDTH)) / _
 MyPrinter.PixPerInchX
MyPrinter.PageHeightInches = CSng(GetDeviceCaps _
 (MyPrinter.Handle, PHYSICALHEIGHT)) / _
 MyPrinter.PixPerInchY
' GetDeviceCaps can get lots of other info which
' I haven't yet bothered with here.
' Set up a default font to Times new Roman size 16
Myfont.lfFaceName = "Tahoma" + Chr$(0)
Myfont.lfEscapement = 0 ' rotation in tenths of a degree
Myfont.lfHeight = 16 * (-MyPrinter.PixPerInchY / _
 POINTSPERINCH) ' 16 point text
Myfont.lfWeight = NORMAL
hNewFont = CreateFontIndirect(Myfont)  'Create the font
' Select our font structure and save previous font info
hOldfont = SelectObject(MyPrinter.Handle, hNewFont)
SetBkMode MyPrinter.Handle, TRANSPARENT ' FontTransparent
' Set the stretch mode
SetStretchBltMode MyPrinter.Handle, STRETCH_DELETESCANS
Exit Sub
dlgError:
UserCancelled = True
End Sub


Private Sub TextPrint(s1 As String, x As Single, y As Single)
Dim xpos As Long, ypos As Long
xpos = x * MyPrinter.PixPerInchX - MyPrinter.OffsetX
ypos = y * MyPrinter.PixPerInchY - MyPrinter.OffsetY
TextOut MyPrinter.Handle, xpos, ypos, s1, Len(s1)
End Sub

Private Sub RectPrint(x1 As Single, y1 As Single, _
 x2 As Single, y2 As Single)
Rectangle MyPrinter.Handle, _
 x1 * MyPrinter.PixPerInchX - MyPrinter.OffsetX, _
 y1 * MyPrinter.PixPerInchY - MyPrinter.OffsetY, _
 x2 * MyPrinter.PixPerInchX - MyPrinter.OffsetX, _
 y2 * MyPrinter.PixPerInchY - MyPrinter.OffsetY
End Sub


Private Function PrintImage(p1 As StdPicture, _
 x As Single, y As Single, _
 wide As Single, high As Single) As Boolean
Dim hOldBitmap As Long, hMemoryDC As Long
hMemoryDC = CreateCompatibleDC(Me.hdc)
hOldBitmap = SelectObject(hMemoryDC, p1)
StretchBlt MyPrinter.Handle, _
 x * MyPrinter.PixPerInchX - MyPrinter.OffsetX, _
 y * MyPrinter.PixPerInchY - MyPrinter.OffsetY, _
 wide * MyPrinter.PixPerInchX, _
 high * MyPrinter.PixPerInchY, _
 hMemoryDC, _
 0, 0, _
 Me.ScaleX(p1.Width, vbHimetric, vbPixels), _
 Me.ScaleY(p1.Height, vbHimetric, vbPixels), _
 vbSrcCopy
hOldBitmap = SelectObject(hMemoryDC, hOldBitmap)
DeleteDC hMemoryDC
End Function


Private Sub Command1_Click()
On Error GoTo errdet:
 If SelectPrinter(Me, Printer.DeviceName) Then
Printer.TrackDefault = False
Printer.ScaleMode = vbUser
End If
Printer.FONTSIZE = 12
If Xunit <> "" Then

Xbound = (Xmax - Xmin) / 20
Ybound = (Ymax - Ymin) / 20


Printer.ScaleMode = ScaleModeConstants.vbUser
Printer.Scale (Xmin - Xbound, (Ymin - Ybound))-(Xmax + Xbound, (Ymax + Ybound))

    'Printer.ScaleTop = Ymin - Ybound
    'Printer.ScaleLeft = Xmin - Xbound
    'Printer.ScaleHeight = Ymax + Ybound
    'Printer.ScaleWidth = Xmax + Xbound

Printer.Line (0, Ymin)-(0, Ymax)
Printer.Line (0, Ymax)-(Xmax, Ymax)
'Printer.Line (0, Ymin)-((0 - Xbound) / 4, Ymin + Ybound)
'Printer.Line (0, Ymin)-((0 + Xbound) / 4, Ymin + Ybound)
'Printer.Line (Xmax, Ymax)-(Xmax - Xbound, (Ymax + (Ybound) / 3))
'Printer.Line (Xmax, Ymax)-(Xmax - Xbound, (Ymax - (Ybound / 3)))



For i = Xmin To Xmax - Xbound Step (Xbound * 2)
Printer.Line (i, Ymax)-(i, (Ymax - (Ybound / 3)))
Printer.CurrentX = i - (Xbound / 4)
Printer.CurrentY = (Ymax + (Ybound / 3))
Printer.Print Format(i, "0.0")
Next

Printer.CurrentX = Xmax - Xbound
Printer.CurrentY = (Ymax - Ybound)
Printer.Print Xunit


For i = Ymin To Ymax - Ybound Step (Ybound * 2)
Printer.Line (Xmin, i)-((Xmin + (Xbound / 4)), i)
Printer.CurrentX = Xmin - Xbound
Printer.CurrentY = Ymax + Ymin - i
Printer.Print Format(i, "0.000")
Next

Printer.CurrentX = Xmin + Xbound
Printer.CurrentY = (Ymin - Ybound)
Printer.Print Yunit

End If

Printer.CurrentX = 0
Printer.CurrentY = Ymax
For m = 1 To Ti
Printer.DrawWidth = 1
Printer.Line -(time_plot(m), (Ymax - (Ybound * 2)) + Ymin - data_plot(m)), vbBlue
Next


Printer.FONTSIZE = 8

For p = 1 To peak_count
peak_area(p) = integrate(peak_start(p), peak_end(p))
peak_height(p) = findMax(peak_start(p), peak_end(p))
peak_pos(p) = findmaxpos(peak_start(p), peak_end(p))


'Printer.CurrentX = peak_pos(p)
'Printer.CurrentY = Ymax - peak_height(p)
'Printer.Print "#" + Str(p) + " = " + Str(peak_pos(p))
RotatePrint "<---#" + Trim(Str(p)) + " (" + Trim(Str(peak_pos(p))) + ")", peak_pos(p), Ymax - (Ybound * 1.5) - peak_height(p), 90
Next
Printer.EndDoc
Exit Sub
errdet:
End Sub

Private Sub Command2_Click()
On Error GoTo errdet:
 If SelectPrinter(Me, Printer.DeviceName) Then
Printer.TrackDefault = False
Printer.ScaleMode = vbUser
End If
Printer.FONTSIZE = 10

If Xunit <> "" Then

Xbound = (Xmax - Xmin) / 20
Ybound = (Ymax - Ymin) / 20


Printer.ScaleMode = ScaleModeConstants.vbUser
Printer.Scale (Xmin - Xbound, (Ymin - Ybound))-(Xmax + Xbound, (Ymax + Ybound) * 2)

    'Printer.ScaleTop = Ymin - Ybound
    'Printer.ScaleLeft = Xmin - Xbound
    'Printer.ScaleHeight = Ymax + Ybound
    'Printer.ScaleWidth = Xmax + Xbound

Printer.Line (0, Ymin)-(0, Ymax)
Printer.Line (0, Ymax)-(Xmax, Ymax)
'Printer.Line (0, Ymin)-((0 - Xbound) / 4, Ymin + Ybound)
'Printer.Line (0, Ymin)-((0 + Xbound) / 4, Ymin + Ybound)
'Printer.Line (Xmax, Ymax)-(Xmax - Xbound, (Ymax + (Ybound) / 3))
'Printer.Line (Xmax, Ymax)-(Xmax - Xbound, (Ymax - (Ybound / 3)))


For i = Xmin To Xmax Step (Xbound * 2)
Printer.Line (i, Ymax)-(i, (Ymax - (Ybound / 3)))
Printer.CurrentX = i - (Xbound / 4)
Printer.CurrentY = (Ymax + (Ybound / 3))
Printer.Print Format(i, "0.0")
Next

Printer.CurrentX = Xmax - Xbound
Printer.CurrentY = (Ymax - Ybound)
Printer.Print Xunit


For i = Ymin To Ymax Step (Ybound * 2)
Printer.Line (Xmin, i)-((Xmin + (Xbound / 4)), i)
Printer.CurrentX = Xmin - Xbound
Printer.CurrentY = Ymax + Ymin - i
Printer.Print Format(i, "0.000")
Next

Printer.CurrentX = Xmin + Xbound
Printer.CurrentY = (Ymin - Ybound)
Printer.Print Yunit

End If

Printer.CurrentX = 0
Printer.CurrentY = Ymax
For m = 1 To Ti
Printer.DrawWidth = 1
Printer.Line -(time_plot(m), (Ymax - (Ybound * 2)) + Ymin - data_plot(m)), vbBlue
Next


Printer.FONTSIZE = 8

For p = 1 To peak_count
peak_area(p) = integrate(peak_start(p), peak_end(p))
peak_height(p) = findMax(peak_start(p), peak_end(p))
peak_pos(p) = findmaxpos(peak_start(p), peak_end(p))


'Printer.CurrentX = peak_pos(p)
'Printer.CurrentY = Ymax - peak_height(p)
'Printer.Print "#" + Str(p) + " = " + Str(peak_pos(p))
RotatePrint "<---#" + Trim(Str(p)) + " (" + Trim(Str(peak_pos(p))) + ")", peak_pos(p), Ymax - (Ybound * 1.5) - peak_height(p), 90
Next

Printer.FONTSIZE = 12

Printer.CurrentX = (Xmax + Xbound) / 3
Printer.CurrentY = Ymax + (Ybound * 3)
Printer.Print "รายงานผลการวิเคราะห์"
Printer.Print " "
Printer.FONTSIZE = 10

Printer.Print vbTab + "วิธีวิเคราะห์ : " + Analysis_title
Printer.Print vbTab + "เครื่องตรวจวัด : " + Detector
Printer.Print vbTab + "หน่วยของการตรวจวัด : " + Yunit
Printer.Print vbTab + "หน่วยของเวลา : " + Xunit
Printer.Print vbTab + "ช่วงเวลาเก็บข้อมูล : " + Str(Runtime) + " " + Xunit
Printer.Print vbTab + "-----------------------------------------------------------------------"
Printer.Print vbTab + "| # " + vbTab + "| ตำแหน่ง" + vbTab + "| เริ่ม" + vbTab + "| สิ้นสุด" + vbTab + "| พื้นที่" + vbTab + "|  สูง" + vbTab + "|"
Printer.Print vbTab + "-----------------------------------------------------------------------"

For i = 1 To frmResult.ListView1.ListItems.Count
    With frmResult.ListView1.ListItems(i)
    Printer.Print vbTab + "| " + .Text + vbTab + "| " + Format(Val(.SubItems(1)), "0") + vbTab + vbTab + "| " + Format(Val(.SubItems(2)), "0") + vbTab + "| " + Format(Val(.SubItems(3)), "0") + vbTab + "| " + Format(Val(.SubItems(4)), "0.000") + vbTab + "| " + Format(Val(.SubItems(5)), "0.000") + vbTab + "|"
    
    End With
Next

Printer.EndDoc
Exit Sub
errdet:
End Sub

Private Function SelectPrinter(frmOwner As Form, Optional _
InitialPrinter As String, Optional _
PrintFlags As Long = PD_PRINTSETUP) _
As Boolean
Dim LongPrinterName As String
Dim PrintDlg As PRINTDLG_TYPE
Dim DevMode As DEVMODE_TYPE
Dim DevName As DEVNAMES_TYPE
Dim lpDevMode As Long, lpDevName As Long
Dim bReturn As Integer, OriginalPrinter As String
Dim p1 As Printer, NewPrinterName As String
PrintDlg.lStructSize = Len(PrintDlg)
PrintDlg.hwndOwner = frmOwner.hwnd
PrintDlg.flags = PrintFlags
On Error Resume Next
OriginalPrinter = Printer.DeviceName
If Len(InitialPrinter) > 0 Then
For Each p1 In Printers
If InStr(1, p1.DeviceName, InitialPrinter, _
vbTextCompare) > 0 Then
Set Printer = p1
Exit For
End If
Next
End If
DevMode.dmDeviceName = Printer.DeviceName
DevMode.dmSize = Len(DevMode)
DevMode.dmFields = DM_ORIENTATION
DevMode.dmPaperWidth = Printer.Width
DevMode.dmOrientation = Printer.Orientation
DevMode.dmPaperSize = Printer.PaperSize
On Error GoTo 0
PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or _
GMEM_ZEROINIT, Len(DevMode))
lpDevMode = GlobalLock(PrintDlg.hDevMode)
If lpDevMode > 0 Then
CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
bReturn = GlobalUnlock(PrintDlg.hDevMode)
End If
With DevName
.wDriverOffset = 8
.wDeviceOffset = .wDriverOffset + 1 + Len _
(Printer.DriverName)
.wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
.wDefault = 0
End With
With Printer
DevName.extra = .DriverName & Chr(0) & _
.DeviceName & Chr(0) & .Port & Chr(0)
End With
PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or _
GMEM_ZEROINIT, Len(DevName))
lpDevName = GlobalLock(PrintDlg.hDevNames)
If lpDevName > 0 Then
CopyMemory ByVal lpDevName, DevName, Len(DevName)
bReturn = GlobalUnlock(lpDevName)
End If
If PrintDialog(PrintDlg) <> 0 Then
'
' Mike's amendment to handle long printer names
CopyMemory DevName, ByVal lpDevName, Len(DevName)
LongPrinterName = Mid$(DevName.extra, _
DevName.wDeviceOffset - DevName.wDriverOffset + 1)
LongPrinterName = Left$(LongPrinterName, _
InStr(LongPrinterName, Chr$(0)) - 1)
DoEvents ' allow dialog to remove itself from display
Me.Refresh
SelectPrinter = True
lpDevName = GlobalLock(PrintDlg.hDevNames)
CopyMemory DevName, ByVal lpDevName, 45
bReturn = GlobalUnlock(lpDevName)
GlobalFree PrintDlg.hDevNames
lpDevMode = GlobalLock(PrintDlg.hDevMode)
CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
bReturn = GlobalUnlock(PrintDlg.hDevMode)
GlobalFree PrintDlg.hDevMode
NewPrinterName = UCase$(Left(DevMode.dmDeviceName, _
InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
' Code now handles long printer names properly
If Printer.DeviceName <> _
LongPrinterName Then
For Each p1 In Printers
If p1.DeviceName = _
LongPrinterName Then
Set Printer = p1
End If
Next
End If
On Error Resume Next
' Transfer settings from the Devmode structure to the
' VB printer object (this example just transfers some
' of them but you can of course use transfer more)
Printer.Copies = DevMode.dmCopies
Printer.Duplex = DevMode.dmDuplex
Printer.Orientation = DevMode.dmOrientation
Printer.PaperSize = DevMode.dmPaperSize
Printer.PrintQuality = DevMode.dmPrintQuality
Printer.ColorMode = DevMode.dmColor
Printer.PaperBin = DevMode.dmDefaultSource
SetBkMode Printer.hdc, TRANSPARENT
'
On Error GoTo 0
Else
SelectPrinter = False ' user cancelled
For Each p1 In Printers
If p1.DeviceName = OriginalPrinter Then
Set Printer = p1
Exit For
End If
Next
GlobalFree PrintDlg.hDevNames
GlobalFree PrintDlg.hDevMode
End If
End Function

Private Sub Command3_Click()
On Error GoTo errdet:
 If SelectPrinter(Me, Printer.DeviceName) Then
Printer.TrackDefault = False
Printer.ScaleMode = vbUser
End If
Printer.FONTSIZE = 12
Printer.CurrentX = Xmax / 4
Printer.CurrentY = 0
Printer.Print "รายงานผลการวิเคราะห์"
Printer.Print " "
Printer.FONTSIZE = 10
Printer.Print vbTab + "วิธีวิเคราะห์ : " + Analysis_title
Printer.Print vbTab + "เครื่องตรวจวัด : " + Detector
Printer.Print vbTab + "หน่วยของการตรวจวัด : " + Yunit
Printer.Print vbTab + "หน่วยของเวลา : " + Xunit
Printer.Print vbTab + "ช่วงเวลาเก็บข้อมูล : " + Str(Runtime) + " " + Xunit
Printer.Print vbTab + "-----------------------------------------------------------------------"
Printer.Print vbTab + "| # " + vbTab + "| ตำแหน่ง" + vbTab + "| เริ่ม" + vbTab + "| สิ้นสุด" + vbTab + "| พื้นที่" + vbTab + "| สูง" + vbTab + "|"
Printer.Print vbTab + "-----------------------------------------------------------------------"

For i = 1 To frmResult.ListView1.ListItems.Count
    With frmResult.ListView1.ListItems(i)
    Printer.Print vbTab + "| " + .Text + vbTab + "| " + Format(Val(.SubItems(1)), "0") + vbTab + vbTab + "| " + Format(Val(.SubItems(2)), "0") + vbTab + "| " + Format(Val(.SubItems(3)), "0") + vbTab + "| " + Format(Val(.SubItems(4)), "0.000") + vbTab + "| " + Format(Val(.SubItems(5)), "0.000") + vbTab + "|"
    
    End With
Next

Printer.EndDoc

Exit Sub
errdet:
End Sub

Private Sub Form_Load()
mainMDI.Toolbar1.Buttons(5).Value = tbrPressed

End Sub

Private Sub Form_Unload(Cancel As Integer)
mainMDI.Toolbar1.Buttons(5).Value = tbrUnpressed
End Sub

