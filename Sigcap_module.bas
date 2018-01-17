Attribute VB_Name = "Sigcap_module"
Public Xmax, Xmin, Ymax, Ymin, Xbound, Ybound As Single
Public SumX, SumY, SumXY, SumX2, SumY2, RegA, RegB, RSquare
Public cXmin, cXmax, cYmin, cYmax, cXbound, cYbound  As Single
Public cXunit As String
Public Xunit, Yunit As String
Public time_plot() As Single
Public data_plot() As Single
Public TimeScale, Runtime, Current_Value, Zero_Cal As Single
Public Ti As Single
Public FileWasOpen, LastPoint, UseUSB As Boolean
Public peak_height() As Single
Public peak_start() As Single
Public peak_end() As Single
Public peak_area() As Single
Public peak_pos() As Single
Public peak_count, filecount As Integer
Public peak_data, Treshold, Baseline As Single
Public Analysis_title, Detector As String
Public CommReady As Boolean
Public Comm_buffer, My_buffer As String
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As String * 1
    lfUnderline As String * 1
    lfStrikeOut As String * 1
    lfCharSet As String * 1
    lfOutPrecision As String * 1
    lfClipPrecision As String * 1
    lfQuality As String * 1
    lfPitchAndFamily As String * 1
    lfFaceName As String * 32
End Type

Public Sub RotateText(PBCtrl As PictureBox, disptxt As String, CX, CY, degree)
Dim Font As LOGFONT
Dim hFont As Long, ret As Long
Const FONTSIZE = 8  ' Desired point size of font

Font.lfEscapement = degree * 10  ' degree rotation
Font.lfFaceName = "Tahoma" + Chr$(0)
Font.lfWeight = 20

' Windows expects the font size to be in pixels and to be negative if you are specifying the character height you want.

Font.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
hFont = CreateFontIndirect(Font)
SelectObject PBCtrl.hdc, hFont

PBCtrl.CurrentX = CX
PBCtrl.CurrentY = CY
PBCtrl.Print disptxt

' Clean up by restoring original font.
ret = DeleteObject(hFont)
End Sub

Public Sub RotatePrint(disptxt As String, CX, CY, degree)
      Dim lf As LOGFONT
      Dim result As Long
      Dim hOldfont As Long
      Dim hPrintDc As Long
      Dim hFont As Long
      Dim setCorX, setCorY As Long
      Const FONTSIZE = 8  ' Desired point size of font

      hPrintDc = Printer.hdc
      
lf.lfEscapement = degree * 10
lf.lfFaceName = "Tahoma" + Chr$(0)
lf.lfWeight = 20


    setCorX = Printer.ScaleX(CX, vbUser, vbPixels) + 100
    setCorY = Printer.ScaleY(CY, vbUser, vbPixels) + 50

      lf.lfHeight = (DESIREDFONTSIZE * -20) / Printer.TwipsPerPixelY
      hFont = CreateFontIndirect(lf)
      hOldfont = SelectObject(hPrintDc, hFont)
      result = TextOut(hPrintDc, setCorX, setCorY, disptxt, Len(disptxt))
      result = SelectObject(hPrintDc, hOldfont)
    
      'Printer.CurrentX = CX
      'Printer.CurrentY = CY
      
      
      result = DeleteObject(hFont)
End Sub





Sub Main()


frmSplash.Show
End Sub

Public Sub count_peak()
peak_count = 0
For k = 1 To Ti
    peak_data = data_plot(k)
    If peak_data > (Baseline + Treshold) Then
        peak_count = peak_count + 1
        ReDim Preserve peak_start(peak_count)
        ReDim Preserve peak_end(peak_count)
        ReDim Preserve peak_height(peak_count)
        ReDim Preserve peak_area(peak_count)
        ReDim Preserve peak_pos(peak_count)

        peak_start(peak_count) = k - 1
    For l = k To Ti
            peak_data = data_plot(l)
            If peak_data < (Baseline) Then
            peak_end(peak_count) = l + 1
            k = l
            Exit For
            End If
    Next l
    End If
Next k

End Sub

Public Function findMax(peakstart, peakstop) As Single
    findMax = -1000
    For q = peakstart To peakstop
    If findMax < data_plot(q) Then findMax = data_plot(q)
    Next
End Function

Public Function findMin(peakstart, peakstop) As Single
    findMin = 1000000
    For q = peakstart To peakstop
    If findMin > data_plot(q) Then findMin = data_plot(q)
    Next
End Function

Public Function findMode(ByRef pValues() As Single) As Single
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict.CompareMode = BinaryCompare
    Dim i As Long
    Dim pCurVal As Single
    For i = 0 To UBound(pValues)
    'limit the values that have to be analyzed to desired precision'
        pCurVal = Round(pValues(i), 2)
            If (pCurVal > 0) Then
                'this will create a dictionary entry if it doesn't exist
                dict.Item(pCurVal) = dict.Item(pCurVal) + 1
            End If
    Next

    'find index of first largest frequency'
    Dim KeyArray, itemArray
    KeyArray = dict.Keys
    itemArray = dict.Items
    pCount = 0
    Dim pModeIdx As Integer
    'find index of mode'
    For i = 0 To UBound(itemArray)
        If (itemArray(i) > pCount) Then
            pCount = itemArray(i)
            pModeIdx = i
        End If
    Next
    'get value corresponding to selected mode index'
    findMode = KeyArray(pModeIdx)
    Set dict = Nothing

End Function

Public Function findmaxpos(peakstart, peakstop) As Single
    Dim datamax
    datamax = 0
    For q = peakstart To peakstop
    If datamax < data_plot(q) Then
    datamax = data_plot(q)
    findmaxpos = time_plot(q)
    End If
    Next
End Function


Public Function integrate(peakstart, peakstop) As Single
        integrate = 0

    For r = peakstart To peakstop
    integrate = integrate + Abs((data_plot(r) - Baseline))
    Next

End Function

Public Sub graph_redraw()
frmChromatogram.Picture1.DrawWidth = 1
'frmChromatogram.Picture1.FontName = "Tahoma" + Chr(0)
'frmChromatogram.Picture1.FontBold = True
frmChromatogram.Picture1.FONTSIZE = 7
'frmChromatogram.Picture1.FontBold = True
frmChromatogram.Picture1.Width = frmChromatogram.Width - 130
frmChromatogram.Picture1.Height = frmChromatogram.Height - 400
frmChromatogram.Picture1.Cls

If Xunit <> "" Then

Xbound = (Xmax - Xmin) / 20
Ybound = (Ymax - Ymin) / 20
frmChromatogram.Picture1.Scale (Xmin - Xbound, Ymax + Ybound)-(Xmax + Xbound, Ymin - Ybound)

'frmChromatogram.Picture1.Line (0, Ymax)-((0 - Xbound) / 4, Ymax - Ybound)
'frmChromatogram.Picture1.Line (0, Ymax)-((0 + Xbound) / 4, Ymax - Ybound)
'frmChromatogram.Picture1.Line (Xmax, Ymin)-(Xmax - Xbound, (Ymin + (Ybound) / 3))
'frmChromatogram.Picture1.Line (Xmax, Ymin)-(Xmax - Xbound, (Ymin - (Ybound / 3)))



For i = Xmin To Xmax Step Xbound * 2
frmChromatogram.Picture1.ForeColor = &HE0E0E0
frmChromatogram.Picture1.Line (i, Ymin)-(i, Ymax)
frmChromatogram.Picture1.ForeColor = &H0
frmChromatogram.Picture1.Line (i, Ymin)-(i, (Ymin + (Ybound / 3)))
frmChromatogram.Picture1.CurrentX = i - (Xbound / 4)
frmChromatogram.Picture1.CurrentY = (Ymin - (Ybound / 3))
frmChromatogram.Picture1.Print Format(i, "0.0")
Next

frmChromatogram.Picture1.CurrentX = Xmax - Xbound
frmChromatogram.Picture1.CurrentY = (Ymin + Ybound)
frmChromatogram.Picture1.Print Xunit


For i = Ymin To Ymax Step Ybound * 2
frmChromatogram.Picture1.ForeColor = &HE0E0E0
frmChromatogram.Picture1.Line (Xmin, i)-(Xmax, i)
frmChromatogram.Picture1.ForeColor = &H0
frmChromatogram.Picture1.Line (Xmin, i)-((Xmin + Xbound) / 5, i)
frmChromatogram.Picture1.CurrentX = Xmin - Xbound
frmChromatogram.Picture1.CurrentY = i
frmChromatogram.Picture1.Print Format(i, "0.000")
Next
'frmChromatogram.Picture1.Line (Xmin, Ymax)-((Xmin + Xbound) / 5, Ymax)
'frmChromatogram.Picture1.CurrentX = Xmin - Xbound
'frmChromatogram.Picture1.CurrentY = Ymax
'frmChromatogram.Picture1.Print Format(i, "0.000")

frmChromatogram.Picture1.CurrentX = Xmin
frmChromatogram.Picture1.CurrentY = (Ymax + Ybound)
frmChromatogram.Picture1.Print Yunit

End If

frmChromatogram.Picture1.CurrentX = 0
frmChromatogram.Picture1.CurrentY = 0
For m = 1 To Ti
frmChromatogram.Picture1.DrawWidth = 1
frmChromatogram.Picture1.Line -(time_plot(m), data_plot(m)), vbBlue
Next

frmChromatogram.Picture1.Line (0, Ymin)-(0, Ymax)
frmChromatogram.Picture1.Line (0, Ymin)-(Xmax, Ymin)

End Sub


Sub process_peak()
Dim iList As ListItem
frmResult.ListView1.ListItems.Clear
For p = 1 To peak_count
peak_area(p) = integrate(peak_start(p), peak_end(p))
peak_height(p) = findMax(peak_start(p), peak_end(p))
peak_pos(p) = findmaxpos(peak_start(p), peak_end(p))

frmChromatogram.Picture1.DrawWidth = 2
frmChromatogram.Picture1.Line (time_plot(peak_start(p)), data_plot(peak_start(p)))-(time_plot(peak_end(p)), data_plot(peak_end(p))), vbBlack
frmChromatogram.Picture1.DrawWidth = 1

For i = peak_start(p) + 1 To peak_end(p)
frmChromatogram.Picture1.DrawWidth = 2
frmChromatogram.Picture1.Line (time_plot(i), data_plot(i))-(time_plot(i), Baseline), vbGreen
frmChromatogram.Picture1.DrawWidth = 1
Next
RotateText frmChromatogram.Picture1, "#" + Str(p) + "<<" + Str(peak_pos(p)), peak_pos(p), peak_height(p), 90


Set iList = frmResult.ListView1.ListItems.Add(, , "#" + Trim(Str(p)))
With iList

.SubItems(1) = Format(peak_pos(p), "0")
.SubItems(2) = Format(time_plot(peak_start(p)), "0")
.SubItems(3) = Format(time_plot(peak_end(p)), "0")
.SubItems(4) = Format(peak_area(p), "0.000")
.SubItems(5) = Format(peak_height(p), "0.000")

End With
Next
frmResult.Show
End Sub

Sub getVal()
'Current_Value = Val(Format(Rnd * 40, "0.00"))
End Sub

Function mysleep(dwloop As Long)
Dim ii
For ii = 1 To dwloop * 100
DoEvents
Next

End Function

Public Sub Redraw_Calibration()
Dim iX, iY As Single
On Error GoTo errdet:
  frmCalibration.Picture1.DrawWidth = 1
'frmCalibration.Picture1.FontName = "Tahoma" + Chr(0)
frmCalibration.Picture1.FontBold = True
frmCalibration.Picture1.FONTSIZE = 7
'frmCalibration.Picture1.FontBold = True
frmCalibration.Picture1.Width = frmCalibration.Width - 130
frmCalibration.Picture1.Height = frmCalibration.Height - 400
frmCalibration.Picture1.Cls

cXmin = 0
cYmin = 0
cXmax = findcMax
cYmax = findSigMax
cXunit = FrmDataCal.Combo1

If cXunit <> "" Then

cXbound = (cXmax - cXmin) / 10
cYbound = (cYmax - cYmin) / 10
frmCalibration.Picture1.Scale (cXmin - cXbound, cYmax + cYbound)-(cXmax + cXbound, cYmin - cYbound)

'frmCalibration.Picture1.Line (0, Ymax)-((0 - cXBound) / 4, Ymax - cYBound)
'frmCalibration.Picture1.Line (0, Ymax)-((0 + cXBound) / 4, Ymax - cYBound)
'frmCalibration.Picture1.Line (cXmax, cYmin)-(cXmax - cXBound, (cYmin + (cYBound) / 3))
'frmCalibration.Picture1.Line (cXmax, cYmin)-(cXmax - cXBound, (cYmin - (cYBound / 3)))



For i = cXmin To cXmax Step cXbound * 2
frmCalibration.Picture1.ForeColor = &HE0E0E0
frmCalibration.Picture1.Line (i, cYmin)-(i, cYmax)
frmCalibration.Picture1.ForeColor = &H0
frmCalibration.Picture1.Line (i, cYmin)-(i, (cYmin + (cYbound / 3)))
frmCalibration.Picture1.CurrentX = i - (cXbound / 4)
frmCalibration.Picture1.CurrentY = (cYmin - (cYbound / 3))
frmCalibration.Picture1.Print Format(i, "0.00")
Next

frmCalibration.Picture1.CurrentX = cXmax - cXbound
frmCalibration.Picture1.CurrentY = (cYmin + cYbound)
frmCalibration.Picture1.Print cXunit


For i = cYmin To Ymax Step cYbound * 2
frmCalibration.Picture1.ForeColor = &HE0E0E0
frmCalibration.Picture1.Line (cXmin, i)-(cXmax, i)
frmCalibration.Picture1.ForeColor = &H0
frmCalibration.Picture1.Line (cXmin, i)-((cXmin + cXbound) / 5, i)
frmCalibration.Picture1.CurrentX = cXmin - cXbound
frmCalibration.Picture1.CurrentY = i
frmCalibration.Picture1.Print Format(i, "0.000")
Next
'frmCalibration.Picture1.Line (cXmin, Ymax)-((cXmin + cXBound) / 5, Ymax)
'frmCalibration.Picture1.CurrentX = cXmin - cXBound
'frmCalibration.Picture1.CurrentY = Ymax
'frmCalibration.Picture1.Print Format(i, "0.000")

frmCalibration.Picture1.CurrentX = cXmin
frmCalibration.Picture1.CurrentY = (cYmax + cYbound)
frmCalibration.Picture1.Print Yunit

For i = 1 To FrmDataCal.ListView1.ListItems.Count
frmCalibration.Picture1.DrawWidth = 2
frmCalibration.Picture1.ForeColor = vbRed
iX = Val(FrmDataCal.ListView1.ListItems(i).SubItems(1))
iY = Val(FrmDataCal.ListView1.ListItems(i).SubItems(3))
frmCalibration.Picture1.Circle (iX, iY), cXbound / 8

Next

DoRegression
frmCalibration.Picture1.CurrentX = cXbound * 3
frmCalibration.Picture1.CurrentY = cYmax + cYbound
frmCalibration.Picture1.Print "y = " + Format(RegA, "0.0000") + "x" + " + " + Format(RegB, "0.0000") + " ,r = " + Format(RSquare, "0.0000")
frmCalibration.Picture1.DrawWidth = 1
frmCalibration.Picture1.DrawStyle = DrawStyleConstants.vbSolid

frmCalibration.Picture1.Line (0, RegB)-(cXmax, (RegA * cXmax) + RegB), vbBlue

frmCalibration.Picture1.DrawStyle = DrawStyleConstants.vbSolid

frmCalibration.Picture1.Line (0, cYmin)-(0, cYmax), vbBlack
frmCalibration.Picture1.Line (0, cYmin)-(cXmax, cYmin), vbBlack
End If
errdet:
End Sub

Public Function findcMax() As Single
    findcMax = -1000
    For q = 1 To FrmDataCal.ListView1.ListItems.Count
    If findcMax < Val(FrmDataCal.ListView1.ListItems(q).SubItems(1)) Then findcMax = Val(FrmDataCal.ListView1.ListItems(q).SubItems(1))
    Next
End Function

Public Function findSigMax() As Single
    findSigMax = -1000
    For q = 1 To FrmDataCal.ListView1.ListItems.Count
    If findSigMax < Val(FrmDataCal.ListView1.ListItems(q).SubItems(3)) Then findSigMax = Val(FrmDataCal.ListView1.ListItems(q).SubItems(3))
    Next
End Function

Public Sub DoRegression()
 SumX = 0
 SumY = 0
 SumX2 = 0
 SumY2 = 0
 SumXY = 0
 
 For i = 1 To FrmDataCal.ListView1.ListItems.Count
 SumX = SumX + Val(FrmDataCal.ListView1.ListItems(i).SubItems(1))
 SumX2 = SumX2 + (Val(FrmDataCal.ListView1.ListItems(i).SubItems(1)) ^ 2)
  SumY = SumY + Val(FrmDataCal.ListView1.ListItems(i).SubItems(3))
 SumY2 = SumY2 + (Val(FrmDataCal.ListView1.ListItems(i).SubItems(3)) ^ 2)
SumXY = SumXY + (Val(FrmDataCal.ListView1.ListItems(i).SubItems(1)) * Val(FrmDataCal.ListView1.ListItems(i).SubItems(3)))
 Next
 
 RegA = ((FrmDataCal.ListView1.ListItems.Count * SumXY) - (SumX * SumY)) / ((FrmDataCal.ListView1.ListItems.Count * SumX2) - (SumX ^ 2))
 RegB = (SumY - (RegA * SumX)) / FrmDataCal.ListView1.ListItems.Count
 RSquare = ((FrmDataCal.ListView1.ListItems.Count * SumXY) - (SumX * SumY)) / (Sqr((FrmDataCal.ListView1.ListItems.Count * SumX2) - SumX ^ 2) * Sqr((FrmDataCal.ListView1.ListItems.Count * SumY2) - SumY ^ 2))

End Sub
