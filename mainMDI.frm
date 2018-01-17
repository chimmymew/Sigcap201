VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.MDIForm mainMDI 
   BackColor       =   &H8000000C&
   Caption         =   "สัญญาณวิเคราะห์ ๒.๐๑"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   10710
   Icon            =   "mainMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   1535
      ButtonWidth     =   1376
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "สร้างใหม่"
            Object.ToolTipText     =   "สร้างการวิเคราะห์ใหม่"
            Object.Tag             =   "สร้างการวิเคราะห์ใหม่"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "เปิด"
            Object.ToolTipText     =   "เปิดแฟ้มข้อมูลเก่า"
            Object.Tag             =   "เปิดแฟ้มข้อมูลเก่า"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "บันทึก"
            Object.ToolTipText     =   "บันทึกแฟ้มข้อมูล"
            Object.Tag             =   "บันทึกแฟ้มข้อมูล"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "วิเคราะห์"
            Object.ToolTipText     =   "วิเคราะห์สัญญาณ"
            Object.Tag             =   "วิเคราะห์สัญญาณ"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "พิมพ์"
            Object.ToolTipText     =   "พิมพ์ผลการทดลอง"
            Object.Tag             =   "พิมพ์ผลการทดลอง"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "เชื่อมต่อ"
            Object.ToolTipText     =   "ตั้งค่าช่องเชื่อมต่อ"
            Object.Tag             =   "ตั้งค่าช่องเชื่อมต่อ"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "สัญญาณ"
            Object.ToolTipText     =   "รับสัญญาณเข้า"
            Object.Tag             =   "รับสัญญาณเข้า"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "กราฟ"
            Object.ToolTipText     =   "ตัวควบคุมกราฟ"
            Object.Tag             =   "ตัวควบคุมกราฟ"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ปั๊ม"
            Object.ToolTipText     =   "ตัวควบคุมปั๊ม"
            Object.Tag             =   "ตัวควบคุมปั๊ม"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "แหล่งแสง"
            Object.ToolTipText     =   "ตัวควบคุมแหล่งแสง"
            Object.Tag             =   "ตัวควบคุมแหล่งแสง"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "โปรแกรม"
            Description     =   "โปรแกรม"
            Object.ToolTipText     =   "สร้างโปรแกรมควบคุมแบบอัตโนมัติ"
            Object.Tag             =   "สร้างโปรแกรมควบคุมแบบอัตโนมัติ"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ออก"
            Object.Tag             =   "ออกจากโปรแกรม"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":6852
            Key             =   "สร้างใหม่"
            Object.Tag             =   "สร้างการวิเคราะห์ใหม่"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":6CA4
            Key             =   "เปิด"
            Object.Tag             =   "เปิดแฟ้มข้อมูลเก่า"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":70F6
            Key             =   "บันทึก"
            Object.Tag             =   "บันทึกแฟ้มข้อมูล"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":7548
            Key             =   "วิเคราะห์"
            Object.Tag             =   "วิเคราะห์ผลการทดลอง"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":799A
            Key             =   "พิมพ์"
            Object.Tag             =   "พิมพ์กราฟและผลการวิเคราะห์"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":7DEC
            Key             =   "การเชื่อมต่อ"
            Object.Tag             =   "การเชื่อมต่อ"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":823E
            Key             =   "กราฟของสัญญาณ"
            Object.Tag             =   "กราฟของสัญญาณ"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":8690
            Key             =   "ตั้งค่ากราฟ"
            Object.Tag             =   "ตั้งค่ากราฟ"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":8AE2
            Key             =   "ตัวควบคุมปั๊ม"
            Object.Tag             =   "ตัวควบคุมปั๊ม"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":8F34
            Key             =   "แหล่งกำเนิดแสง"
            Object.Tag             =   "การควบคุมแหล่งแสง"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":9386
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainMDI.frx":97D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2280
      Top             =   360
   End
   Begin MSCommLib.MSComm mCom 
      Left            =   1200
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   1
      InputLen        =   1
      OutBufferSize   =   1
      RThreshold      =   1
      BaudRate        =   19200
      EOFEnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.csv"
      DialogTitle     =   "เปิดแฟ้มสัญญาณ"
      Filter          =   "*.csv"
   End
   Begin VB.Menu fileMenu 
      Caption         =   "แฟ้มข้อมูล"
      Begin VB.Menu SessionMenu 
         Caption         =   "กระบวนการวิเคราะห์"
         Begin VB.Menu SessionParamMenu 
            Caption         =   "พารามิเตอร์"
         End
         Begin VB.Menu sep2 
            Caption         =   "-"
         End
         Begin VB.Menu NewSessionMenu 
            Caption         =   "สร้างใหม่"
         End
         Begin VB.Menu OpenSessionMenu 
            Caption         =   "เรียกจากไฟล์"
         End
         Begin VB.Menu SaveSessionMenu 
            Caption         =   "บันทึกลงไฟล์"
         End
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMenu 
         Caption         =   "ออกจากโปรแกรม"
      End
   End
   Begin VB.Menu DataMenu 
      Caption         =   "วิเคราะห์ข้อมูล"
      Begin VB.Menu DataAnalParam 
         Caption         =   "พารามิเตอร์การวิเคราะห์"
      End
      Begin VB.Menu AutoProcess 
         Caption         =   "การวิเคราะห์อัตโนมัติ"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu AnalysisTool 
         Caption         =   "รายงานและการพิมพ์"
      End
   End
   Begin VB.Menu CommunicationMEnu 
      Caption         =   "การสื่อสาร"
      Begin VB.Menu SetADmenu 
         Caption         =   "ตั้งค่าอุปกรณ์รับข้อมูล"
      End
   End
   Begin VB.Menu ChromatogramMenu 
      Caption         =   "กราฟและโครมาโทแกรม"
      Begin VB.Menu AddChromatogramFromCap 
         Caption         =   "รับสัญญาณใหม่"
      End
      Begin VB.Menu addFromSave 
         Caption         =   "เพิ่มจากที่มีอยู่ก่อนแล้ว"
      End
      Begin VB.Menu GControl 
         Caption         =   "ควบคุมกราฟ"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu DelChroma 
         Caption         =   "ลบออกจากกระบวนการ"
      End
   End
   Begin VB.Menu mnuControl 
      Caption         =   "การควบคุม"
      Begin VB.Menu mnuPump 
         Caption         =   "ปั๊มและเครื่องฉีดสาร"
      End
      Begin VB.Menu mnuRGB 
         Caption         =   "สเปกตรัมของ RGB"
      End
      Begin VB.Menu mnuProgram 
         Caption         =   "โปรแกรม"
      End
   End
End
Attribute VB_Name = "mainMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const VendorID = 43690 ' &HAAAA - Replace with your device's
Private Const ProductID = 61186 ' &HEF04 - product and vendor IDs
' read and write buffers
Private Const BufferInSize = 2
Private Const BufferOutSize = 0
Dim BufferIn(0 To BufferInSize) As Byte
Dim BufferOut(0 To BufferOutSize) As Byte
Dim Tr As Byte
Dim Myval As Single


Private Sub RunMenu_Click()

End Sub

Private Sub AddChromatogramFromCap_Click()

frmChromatogram.Show
frmChromatogram.Height = 5000
frmChromatogram.Width = 8000
ReDim time_plot(0)
ReDim data_plot(0)
Ti = 0
graph_redraw
frmInjection.Show

End Sub

Private Sub addFromSave_Click()
On Error GoTo errdet:
cmd1.DialogTitle = "เปิดแฟ้มสัญญาณที่เก็บไว้"
cmd1.DefaultExt = "*.csv"
cmd1.Filter = "*.csv|*.csv"
cmd1.CancelError = True
cmd1.ShowOpen
Dim dat_in As String
Dim fld() As String

Open cmd1.FileName For Input As #1
Line Input #1, dat_in
Line Input #1, dat_in
fld = Split(dat_in, ",")
TimeScale = fld(0)
'Picture1.MousePointer = MousePointerConstants.vbArrow
Ti = 0
While Not EOF(1)
Line Input #1, dat_in
Ti = Ti + 1
ReDim Preserve time_plot(Ti)
ReDim Preserve data_plot(Ti)
fld = Split(dat_in, ",")
time_plot(Ti) = Val(fld(0))
data_plot(Ti) = Val(fld(1))
Wend

Close #1
graph_redraw
Exit Sub
errdet:
Close #1
MsgBox "เกิดข้อผิดพลาด ไม่สามารถเปิดแฟ้มได้", vbCritical, "ไม่สามารถเปิดแฟ้มได้"
End Sub

Private Sub AnalysisTool_Click()
frmPrint.Show
End Sub

Private Sub DataAnalParam_Click()
frmAnalysis.Show
End Sub

Private Sub DelChroma_Click()
On Error GoTo errdet:
ReDim time_plot(0)
ReDim data_plot(0)
Ti = 0
graph_redraw
Exit Sub
errdet:
End Sub

Private Sub ExitMenu_Click()
MDIForm_QueryUnload 0, 0
End Sub

Private Sub GControl_Click()
frmGraphControl.Show
End Sub

Private Sub mCom_OnComm()


Select Case mCom.CommEvent


       Case comEvReceive
                  
           My_buffer = My_buffer + mCom.Input
           If InStr(1, My_buffer, "*") > 0 Or InStr(1, My_buffer, ">") > 0 Then
           Comm_buffer = My_buffer
              Debug.Print Comm_buffer;
               My_buffer = ""

               CommReady = True
              

              End If
              
       Case comEvSend:  'here put your condition that you want
       Case comEvCTS
       Case comEvDSR
       Case comEvCD
       Case comEvRing
       Case comEvEOF
       Case comBreak
       Case comCDTO
       Case comCTSTO
       Case comDCB
       Case comDSRTO
       Case comFrame
       Case comOverrun
       Case comRxOver
       Case comRxParity
       Case comTxFull
End Select
End Sub

Private Sub MDIForm_Load()
Analysis_title = "การวิเคราะห์ที่ตั้งไว้"
Detector = "ศักย์ไฟฟ้าทั่วไป"
Xunit = "วินาที"
Yunit = "mV"
Xmin = 0
Xmax = 10
TimeScale = 0.02
Runtime = 10
Ymin = 0
Ymax = 100
  'ConnectToHID (Me.hwnd)
   frmMsg.Show
   frmValue.Show
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Msg = "คุณต้องการออกจากโปรแกรมสัญญาณวิเคราะห์จริงหรือ?"
Style = vbYesNo + vbExclamation
Title = "ออกจากโปรแกรม"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
   'DisconnectFromHID

End
ElseIf Response = vbNo Then
Cancel = True
End If
End Sub

Private Sub MDIForm_Terminate()
MDIForm_QueryUnload 0, 0
End Sub

Private Sub mnuProgram_Click()
frmProgram.Show
End Sub

Private Sub mnuPump_Click()
frmPump.Show
End Sub

Private Sub mnuRGB_Click()
frmRGB.Show
End Sub

Private Sub NewSessionMenu_Click()
frmSession.Show
End Sub

Private Sub OpenSessionMenu_Click()
On Error GoTo errdet:

frmChromatogram.Show
frmChromatogram.Height = 5000
frmChromatogram.Width = 8000

cmd1.DialogTitle = "เปิดแฟ้มสัญญาณที่เก็บไว้"
cmd1.Filter = "Signal file (*.signal)|*.signal"
cmd1.CancelError = True
cmd1.ShowOpen
Dim dat_in As String
Dim fld() As String

Open cmd1.FileName For Input As #1
Line Input #1, Analysis_title
Line Input #1, Detector
Line Input #1, Yunit
Line Input #1, Xunit
Line Input #1, dat_in: Runtime = Val(dat_in)
Line Input #1, dat_in: TimeScale = Val(dat_in)
Line Input #1, dat_in: Xmax = Val(dat_in)
Line Input #1, dat_in: Xmin = Val(dat_in)
Line Input #1, dat_in: Ymax = Val(dat_in)
Line Input #1, dat_in: Ymin = Val(dat_in)



'Picture1.MousePointer = MousePointerConstants.vbArrow
Ti = 0
While Not EOF(1)
Line Input #1, dat_in
Ti = Ti + 1
ReDim Preserve time_plot(Ti)
ReDim Preserve data_plot(Ti)
fld = Split(dat_in, ",")
time_plot(Ti) = Val(fld(0))
data_plot(Ti) = Val(fld(1))
Wend

Close #1
graph_redraw


errdet:
Close #1
End Sub

Private Sub SaveSessionMenu_Click()
filecount = filecount + 1
On Error GoTo errdet:
cmd1.DialogTitle = "บันทึกลงไฟล์"
cmd1.Filter = "Signal File(*.signal)|*.signal"
cmd1.FileName = "signal" + Format(filecount, "000")
cmd1.ShowSave

Open cmd1.FileName For Output As #1
Print #1, Analysis_title
Print #1, Detector
Print #1, Yunit
Print #1, Xunit
Print #1, Runtime
Print #1, TimeScale
Print #1, Xmax
Print #1, Xmin
Print #1, Ymax
Print #1, Ymin

For i = 1 To Ti
Print #1, Str(time_plot(i)) + "," + Str(data_plot(i))
Next

Close #1
Exit Sub
errdet:
Close #1
End Sub

Private Sub SessionParamMenu_Click()
frmSession.Show
End Sub

Public Sub OnPlugged(ByVal pHandle As Long)
   If hidGetVendorID(pHandle) = VendorID And hidGetProductID(pHandle) = ProductID Then
      ' ** YOUR CODE HERE **
frmMsg.Label1.Caption = "พบการ์ดรับสัญญาณ!"
frmMsg.Show
frmValue.Show
   End If
End Sub

'*****************************************************************
' a HID device has been unplugged...
'*****************************************************************
Public Sub OnUnplugged(ByVal pHandle As Long)
   If hidGetVendorID(pHandle) = VendorID And hidGetProductID(pHandle) = ProductID Then
      ' ** YOUR CODE HERE **
frmMsg.Label1.Caption = "การ์ดถูกถอดออก!"
frmMsg.Show
frmValue.Label1.Caption = "ไม่พบการ์ด"
frmValue.Show
   End If
End Sub

'*****************************************************************
' controller changed notification - called
' after ALL HID devices are plugged or unplugged

'*****************************************************************
Public Sub OnChanged()
   Dim DeviceHandle As Long
   ' get the handle of the device we are interested in, then set
   ' its read notify flag to true - this ensures you get a read
   ' notification message when there is some data to read...
   DeviceHandle = hidGetHandle(VendorID, ProductID)
   hidSetReadNotify DeviceHandle, True
End Sub

'*****************************************************************
' on read event...

'*****************************************************************
Public Sub OnRead(ByVal pHandle As Long)
 Dim AD_Read As Single

   If hidRead(pHandle, BufferIn(0)) Then
      ' ** YOUR CODE HERE **
            '######################################################################################
'CALCULATION OF TEMPERATURE OF TWO received bytes and display the required format

            AD_Read = (BufferIn(1) * 8) + BufferIn(2)
                If AD_Read > 0 Then Current_Value = AD_Read / 10
                frmValue.Label1.Caption = Format(Current_Value, "0.0") + " " + Yunit
            '######################################################################################

   End If
End Sub

Private Sub SetADmenu_Click()
frmComset.Show
End Sub

Private Sub Timer1_Timer()

       Dim Comm_fld() As String
       
If mCom.PortOpen = True Then
               CommReady = False
                 
                 mCom.Output = "A"
                 
                 Do
                 DoEvents
                 Loop While CommReady = False
            
            Comm_fld = Split(Comm_buffer, ",")
            
            If UBound(Comm_fld) > 1 Then
            If Comm_fld(1) = "A" Then
            Myval = 1.024 - Val(Comm_fld(0))
            Current_Value = Myval - Zero_Cal
            frmValue.Label1.Caption = Format(Current_Value, "0.0000") + " " + Yunit
            End If
       
            End If
            
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
        Case 1
        frmSession.Show
        Button.Value = tbrUnpressed
                
        Case 2
       OpenSessionMenu_Click
        Button.Value = tbrUnpressed
                
         Case 3
         SaveSessionMenu_Click
             Button.Value = tbrUnpressed
             
          Case 4
          DataAnalParam_Click
          
          Case 5
          AnalysisTool_Click
          
          Case 6
          SetADmenu_Click
          
        Case 7
         AddChromatogramFromCap_Click
         
         
        Case 8
        GControl_Click
        
        
        Case 9
        mnuPump_Click
        
        Case 10
        mnuRGB_Click
        
        Case 11
        mnuProgram_Click
        
        Case 12
        MDIForm_Terminate
           
End Select

End Sub
