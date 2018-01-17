VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmProgram 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "โปรแกรมควบคุมแบบอัตโนมัติ"
   ClientHeight    =   7440
   ClientLeft      =   12105
   ClientTop       =   1125
   ClientWidth     =   6480
   Icon            =   "frmProgram.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   6480
   Begin VB.TextBox txt_Program 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1080
      Width           =   6375
   End
   Begin MSComDlg.CommonDialog program_dialog 
      Left            =   5520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "program1"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   7050
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1693
            MinWidth        =   1693
            Text            =   "กำลังทำ"
            TextSave        =   "กำลังทำ"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4303
            MinWidth        =   4303
            Text            =   "คำสั่ง"
            TextSave        =   "คำสั่ง"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "พารามิเตอร์"
            TextSave        =   "พารามิเตอร์"
         EndProperty
      EndProperty
   End
   Begin VB.Timer timer_Program 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   600
   End
   Begin VB.CommandButton cmd_Stop 
      Caption         =   "หยุด"
      Height          =   855
      Left            =   3960
      Picture         =   "frmProgram.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmd_Go 
      Caption         =   "ปฏิบัติ"
      Height          =   855
      Left            =   3000
      Picture         =   "frmProgram.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmd_check 
      Caption         =   "ตรวจ"
      Height          =   855
      Left            =   2040
      Picture         =   "frmProgram.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "บันทึก"
      Height          =   855
      Left            =   1080
      Picture         =   "frmProgram.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmd_Open 
      Caption         =   "เปิด"
      Height          =   855
      Left            =   120
      Picture         =   "frmProgram.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time_elapse As Long
Dim schema_line() As String
Dim schema_command() As String
Dim No_error As Integer
Dim error_Checked As Boolean
Dim pump1_Speed, pump2_Speed As Integer
Dim Mystring As String
Dim Time_wait As Boolean
Dim Wait_Time As Integer
Dim m_Stop As Boolean



Private Sub cmd_check_Click()
check_schema
error_Checked = True
End Sub

Private Sub cmd_Go_Click()
On Error GoTo errdet:
   m_Stop = False
    If error_Checked = False Then
      MsgBox "โปรแกรมไม่ได้ผ่านการตรวจสอบ โปรดตรวจสอบโปรแกรมเสียก่อน", vbInformation, "ไม่ผ่านเงื่อนไขระบบอัตโนมัติ"
    Exit Sub
    End If
    
        If No_error > 0 Then
      MsgBox "ยังมีข้อผิดพลาดบางประการในโปรแกรม โปรดตรวจสอบ", vbInformation, "ไม่ผ่านเงื่อนไขระบบอัตโนมัติ"
    Exit Sub
    End If
    
      Dim schema_no As Integer
    Dim Search, Where, Schema_Process
    
    No_error = 0
    
    schema_line = Split(txt_Program, vbCrLf)
     
     For schema_no = 0 To UBound(schema_line) - 1
        txt_Program.SetFocus
        Search = schema_line(schema_no)
        Where = Schema_Process + Len(schema_line(schema_no))
  
        schema_command = Split(schema_line(schema_no), " ")
       If UBound(schema_command) > 0 Then
       do_Schema schema_command(0), schema_command(1)
       Else
        do_Schema schema_command(0), " "
        End If
        
              txt_Program.SelStart = Where   ' set selection start and
       txt_Program.SelLength = Len(Search)   ' set selection length.
       StatusBar1.Panels(1).Text = "ทำ : " + Str(schema_no + 1)
       StatusBar1.Panels(2).Text = "คำสั่ง : " + schema_command(0)
       StatusBar1.Panels(3).Text = ""
        If UBound(schema_command) > 0 Then StatusBar1.Panels(3).Text = "พารามิเตอร์ : " + schema_command(1)
        If UBound(schema_command) > 1 Then StatusBar1.Panels(3).Text = "พารามิเตอร์ : " + schema_command(1) + " " + schema_command(2)
        
        Schema_Process = Schema_Process + Len(schema_line(schema_no)) + 2

       While Time_wait = True
       DoEvents
       Wend
       If m_Stop = True Then
      MsgBox "หยุดกระบวนการทั้งหมด", vbInformation, "หยุดกระบวนการ"
      do_Schema "ปิดปั๊มแรก", " "
       do_Schema "ปิดปั๊มสอง", " "
       Exit Sub
       End If
     Next
    
    
    
         Exit Sub
errdet:
        MsgBox Err.Description, vbInformation, "เกิดข้อผิดพลาด"
End Sub

Private Sub cmd_Open_Click()
Dim Program_data As String
On Error GoTo errdet:
error_Checked = False
program_dialog.DialogTitle = "เปิดโปรแกรมที่บันทึกไว้"
program_dialog.Filter = "Schema file (*.prog)|*.prog"
program_dialog.CancelError = True
program_dialog.ShowOpen
txt_Program.Text = ""
Open program_dialog.FileName For Input As #3

While Not EOF(3)
Line Input #3, Program_data
txt_Program.Text = txt_Program.Text + Program_data + vbCrLf
Wend

Close #3
Exit Sub
errdet:
Close #3
MsgBox Err.Description, vbInformation, "ผิดพลาด"
End Sub

Private Sub cmd_save_Click()
Dim Program_data As String
On Error GoTo errdet:
program_dialog.DialogTitle = "บันทึกโปรแกรม"
program_dialog.Filter = "Schema file (*.prog)|*.prog"
program_dialog.CancelError = True
program_dialog.ShowSave

Open program_dialog.FileName For Output As #2

Print #2, txt_Program.Text

Close #2
Exit Sub
errdet:
Close #2
MsgBox Err.Description, vbInformation, "ผิดพลาด"
End Sub

Sub check_schema()
On Error GoTo errdet:
    Dim schema_no As Integer
    Dim Search, Where, Schema_Process
    
    No_error = 0
    
    schema_line = Split(txt_Program, vbCrLf)
     
     For schema_no = 0 To UBound(schema_line) - 1
        txt_Program.SetFocus
        Search = schema_line(schema_no)
        Where = Schema_Process + Len(schema_line(schema_no))
  
        schema_command = Split(schema_line(schema_no), " ")
        find_Schema_Error schema_command(0)
              txt_Program.SelStart = Where   ' set selection start and
       txt_Program.SelLength = Len(Search)   ' set selection length.
       StatusBar1.Panels(1).Text = "ตรวจ : " + Str(schema_no + 1)
       StatusBar1.Panels(2).Text = "คำสั่ง : " + schema_command(0)
       StatusBar1.Panels(3).Text = ""

       If UBound(schema_command) > 0 Then StatusBar1.Panels(3).Text = "พารามิเตอร์ : " + schema_command(1)
        If UBound(schema_command) > 1 Then StatusBar1.Panels(3).Text = "พารามิเตอร์ : " + schema_command(1) + " " + schema_command(2)

       DoEvents
       Sleep 100
        Schema_Process = Schema_Process + Len(schema_line(schema_no)) + 2
     Next
        
        MsgBox "ตรวจเสร็จเรียบร้อยแล้ว พบข้อผิดพลาด จำนวน" + Str(No_error) + " คำสั่ง", vbInformation, "ตรวจคำสั่ง"
        
        Exit Sub
errdet:
        MsgBox Err.Description, vbInformation, "เกิดข้อผิดพลาด"
        error_Checked = False
End Sub

Sub find_Schema_Error(schema As String)
   Select Case schema
   Case ""
   Case "เริ่มระบบควบคุมอัตโนมัติ"
   Case "ตั้งค่าการวิเคราะห์"
   Case "ตั้งเวลาการบันทึก"
   Case "ใช้ตัวตรวจวัด"
   Case "บันทึกสัญญาณทุก"
   Case "หน่วยที่ใช้วัด"
   Case "หน่วยของเวลา"
   Case "ตั้งความเร็วปั๊มแรก"
    Case "เปิดปั๊มแรก"
     Case "ปิดปั๊มแรก"
     Case "เปลี่ยนความเร็วปั๊มแรก"
     Case "ตั้งความเร็วปั๊มสอง"
    Case "เปิดปั๊มสอง"
     Case "ปิดปั๊มสอง"
      Case "เปลี่ยนความเร็วปั๊มสอง"
    Case "รอคำสั่งถัดไป"
    Case "เริ่มการบันทึกสัญญาณ"
    Case "รอจนหมดเวลาบันทึก"
     Case "บันทึกสัญญาณชื่อ"

   
   Case Else
   MsgBox "ไม่รู้จักคำสั่ง : '" + schema + "'", vbInformation, "การตรวจรายการโปรแกรม"
   No_error = No_error + 1
   End Select
End Sub

Sub do_Schema(schema As String, param As String)
Select Case schema
    Case "เริ่มระบบควบคุมอัตโนมัติ"
   Case "ตั้งค่าการวิเคราะห์"
    Analysis_title = param
    '-----------------------------------------------------------------------------------------------------------------------------
   Case "ตั้งเวลาการบันทึก"
   Runtime = Val(param)
  Xmax = Val(param)
   '-----------------------------------------------------------------------------------------------------------------------------
   Case "ใช้ตัวตรวจวัด"
   Detector = param
   '-----------------------------------------------------------------------------------------------------------------------------
   Case "บันทึกสัญญาณทุก"
   TimeScale = Val(param)
   '-----------------------------------------------------------------------------------------------------------------------------
   Case "หน่วยที่ใช้วัด"
   Yunit = param
   '-----------------------------------------------------------------------------------------------------------------------------
   Case "หน่วยของเวลา"
   Xunit = param
   '-----------------------------------------------------------------------------------------------------------------------------
   Case "ตั้งความเร็วปั๊มแรก"
   pump1_Speed = 9000 + (500 * (1.2 - Val(param)))
'-----------------------------------------------------------------------------------------------------------------------------
   Case "เปลี่ยนความเร็วปั๊มแรก"
   pump1_Speed = 9000 + (500 * (1.2 - Val(param)))
     If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True
mainMDI.Timer1.Enabled = False

Do

DoEvents
Loop While CommReady = False

mainMDI.mCom.Output = "C"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

Mystring = Trim(Str(pump1_Speed))
mainMDI.mCom.Output = Mystring
mainMDI.mCom.Output = Chr(13)


Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0

mainMDI.Timer1.Enabled = True
   
'-----------------------------------------------------------------------------------------------------------------------------


    Case "เปิดปั๊มแรก"
    If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True
mainMDI.Timer1.Enabled = False


mainMDI.mCom.Output = "C"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

Mystring = Trim(Str(pump1_Speed))
mainMDI.mCom.Output = Mystring
mainMDI.mCom.Output = Chr(13)


Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0

mainMDI.Timer1.Enabled = True

'-----------------------------------------------------------------------------------------------------------------------------

     Case "ปิดปั๊มแรก"
   
If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True

mainMDI.Timer1.Enabled = False

mainMDI.mCom.Output = "C"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

Mystring = "10000"
mainMDI.mCom.Output = Mystring
mainMDI.mCom.Output = Chr(13)
Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0
mainMDI.Timer1.Enabled = True
     
 '-----------------------------------------------------------------------------------------------------------------------------
 
     Case "ตั้งความเร็วปั๊มสอง"
      pump2_Speed = 9000 + (500 * (1.2 - Val(param)))
  
     '-----------------------------------------------------------------------------------------------------------------------------
     
    Case "เปิดปั๊มสอง"
    
    If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True
mainMDI.Timer1.Enabled = False



mainMDI.mCom.Output = "I"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

Mystring = Trim(Str(pump2_Speed))
mainMDI.mCom.Output = Mystring
mainMDI.mCom.Output = Chr(13)


Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0

mainMDI.Timer1.Enabled = True
    
  '-----------------------------------------------------------------------------------------------------------------------------
    
    Case "เปลี่ยนความเร็วปั๊มสอง"
   pump2_Speed = 9000 + (500 * (1.2 - Val(param)))
     If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True
mainMDI.Timer1.Enabled = False



mainMDI.mCom.Output = "I"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

Mystring = Trim(Str(pump2_Speed))
mainMDI.mCom.Output = Mystring
mainMDI.mCom.Output = Chr(13)


Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0

mainMDI.Timer1.Enabled = True
   
'-----------------------------------------------------------------------------------------------------------------------------
    
    
     Case "ปิดปั๊มสอง"
     
  
If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True

mainMDI.Timer1.Enabled = False

mainMDI.mCom.Output = "I"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

Mystring = "10000"
mainMDI.mCom.Output = Mystring
mainMDI.mCom.Output = Chr(13)
Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0
mainMDI.Timer1.Enabled = True

'-----------------------------------------------------------------------------------------------------------------------------


     
    Case "รอคำสั่งถัดไป"
    
    Time_wait = True
    Wait_Time = Val(param)
    timer_Program.Enabled = True
    
    '-----------------------------------------------------------------------------------------------------------------------------

    
    Case "เริ่มการบันทึกสัญญาณ"
    frmChromatogram.Show
    frmInjection.Show
    graph_redraw
    
    Select Case Xunit
    Case "วินาที"
    frmInjection.Timer1.Interval = 1000 * TimeScale
     Case "นาที"
    frmInjection.Timer1.Interval = 60000 * TimeScale
End Select

 frmInjection.Command1.Caption = "หยุด"
frmInjection.Timer1.Enabled = True
frmValue.Show

'-----------------------------------------------------------------------------------------------------------------------------
    
    
    Case "รอจนหมดเวลาบันทึก"
     Case "บันทึกสัญญาณชื่อ"
     End Select
End Sub

Private Sub cmd_Stop_Click()
m_Stop = True
End Sub

Private Sub Form_Load()
error_Checked = False
End Sub

Private Sub timer_Program_Timer()
Wait_Time = Wait_Time - 1
StatusBar1.Panels(3).Text = "เวลา :" + Str(Wait_Time) + " วินาที"
If Wait_Time = 0 Then
Time_wait = False
timer_Program.Enabled = False
End If

  If m_Stop = True Then
      MsgBox "หยุดกระบวนการทั้งหมด", vbInformation, "หยุดกระบวนการ"
      do_Schema "ปิดปั๊มแรก", " "
       do_Schema "ปิดปั๊มสอง", " "
       timer_Program.Enabled = False
       Exit Sub
       End If

End Sub
