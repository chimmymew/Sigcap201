VERSION 5.00
Begin VB.Form frmPump 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "แผงควบคุมปั๊ม"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "frmPump.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4305
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmPump.frx":0442
      Left            =   1440
      List            =   "frmPump.frx":0467
      TabIndex        =   9
      Text            =   "--"
      Top             =   600
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ฉีด"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   7
      Top             =   405
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmPump.frx":049D
      Left            =   1440
      List            =   "frmPump.frx":04C2
      TabIndex        =   4
      Text            =   "100"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ปั๊ม"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmPump.frx":04FD
      Left            =   1440
      List            =   "frmPump.frx":0513
      TabIndex        =   0
      Text            =   "1.0"
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "วินาที"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ฉีดอัตโนมัติ ทุก ๆ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "ไมโครลิตร"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ปั๊มฉีดสารตัวอย่าง"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "มิลลิลิตร/นาที"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "อัตราการไหลปั๊ม"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmPump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mystring As String
Dim Ti As Long

Private Sub Command1_Click()
Dim pump_speed As Long

pump_speed = 9000 + (500 * (1.2 - Val(Combo1)))


If Command1.Caption = "ปั๊ม" Then
Command1.Enabled = False
If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True
mainMDI.Timer1.Enabled = False

Do

DoEvents
Loop While CommReady = False

mainMDI.mCom.Output = "C"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

mystring = Trim(Str(pump_speed))
mainMDI.mCom.Output = mystring
mainMDI.mCom.Output = Chr(13)


Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0


Command1.Caption = "หยุด"

mainMDI.Timer1.Enabled = True
Command1.Enabled = True
Else


If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True

mainMDI.Timer1.Enabled = False
Command1.Enabled = False
Do

DoEvents
Loop While CommReady = False

mainMDI.mCom.Output = "C"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

mystring = "10000"
mainMDI.mCom.Output = mystring
mainMDI.mCom.Output = Chr(13)
Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0
Command1.Caption = "ปั๊ม"
mainMDI.Timer1.Enabled = True
Command1.Enabled = True
End If
End Sub

Private Sub Command2_Click()
Dim inject_vol As Integer

If Val(Combo3) <> 0 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
Ti = 0
End If
Command2.Caption = "รอ"
Command2.Enabled = False
inject_vol = Val(Combo2) * 10
Timer1.Interval = inject_vol
If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True
mainMDI.Timer1.Enabled = False

Do

DoEvents
Loop While CommReady = False

mainMDI.mCom.Output = "I"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

mystring = "9000"
mainMDI.mCom.Output = mystring
mainMDI.mCom.Output = Chr(13)

Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0

Timer1.Enabled = True

End Sub

Private Sub Form_Load()
mainMDI.Toolbar1.Buttons(9).Value = tbrPressed
End Sub

Private Sub Form_Unload(Cancel As Integer)
mainMDI.Toolbar1.Buttons(9).Value = tbrUnpressed
End Sub

Private Sub Timer1_Timer()
If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True


Do

DoEvents
Loop While CommReady = False

mainMDI.mCom.Output = "I"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

mystring = "10000"
mainMDI.mCom.Output = mystring
mainMDI.mCom.Output = Chr(13)

Do

DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0

Command2.Caption = "ฉีด"
Command2.Enabled = True
mainMDI.Timer1.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Ti = Ti + 1

If Val(Combo3) <> 0 Then
If Ti Mod Val(Combo3) = 0 Then Command2_Click
Else
Timer2.Enabled = False
End If
End Sub
