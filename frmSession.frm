VERSION 5.00
Begin VB.Form frmSession 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "พารามิเตอร์พื้นฐานสำหรับการวิเคราะห์สัญญาณ"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7500
   Icon            =   "frmSession.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   7500
   Begin VB.CommandButton Command1 
      Caption         =   "ยกเลิก"
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
      Index           =   1
      Left            =   4920
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ตกลง"
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
      Index           =   0
      Left            =   3960
      TabIndex        =   14
      Top             =   1200
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
      Index           =   4
      ItemData        =   "frmSession.frx":0442
      Left            =   6360
      List            =   "frmSession.frx":044F
      TabIndex        =   13
      Text            =   "วินาที"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   12
      Text            =   "1"
      Top             =   840
      Width           =   495
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
      Index           =   3
      ItemData        =   "frmSession.frx":046A
      Left            =   6360
      List            =   "frmSession.frx":0477
      TabIndex        =   10
      Text            =   "วินาที"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   8
      Text            =   "10"
      Top             =   480
      Width           =   495
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
      Index           =   2
      ItemData        =   "frmSession.frx":0492
      Left            =   1800
      List            =   "frmSession.frx":049F
      TabIndex        =   7
      Text            =   "วินาที"
      Top             =   1200
      Width           =   1935
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
      Index           =   1
      ItemData        =   "frmSession.frx":04BA
      Left            =   1800
      List            =   "frmSession.frx":04D6
      TabIndex        =   5
      Text            =   "mABS"
      Top             =   840
      Width           =   1935
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
      Index           =   0
      ItemData        =   "frmSession.frx":04FD
      Left            =   1800
      List            =   "frmSession.frx":050D
      TabIndex        =   3
      Text            =   "ตัวตรวจวัดด้วยแสง"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Text            =   "การวิเคราะห์ที่ตั้งไว้"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "บันทึกสัญญาณทุก ๆ "
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
      Index           =   5
      Left            =   3840
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "เวลาที่ใช้ในการวิเคราะห์"
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
      Index           =   4
      Left            =   3840
      TabIndex        =   9
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "หน่วยของเวลา"
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
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "หน่วยที่ใช้วัด"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "เครื่องตรวจวัดเป็น"
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
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "กระบวนการนี้เป็นการวิเคราะห์เกี่ยวกับ"
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
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change(Index As Integer)
Select Case Index
Case 2: Combo1(3) = Combo1(2)
Combo1(4) = Combo1(2)
End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
Select Case Index
Case 2: Combo1(3) = Combo1(2)
Combo1(4) = Combo1(2)
End Select

End Sub

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
Analysis_title = Text1(0)
Detector = Combo1(0)
Yunit = Combo1(1)
Xunit = Combo1(2)
Runtime = Val(Text1(1))
TimeScale = Val(Text1(2))
Xmax = Val(Text1(1))
'Ymax = 50
'Xmin = 0
'Ymin = 0
End If
Unload Me
End Sub


Private Sub Form_Load()
If Analysis_title <> "" Then
Text1(0) = Analysis_title
Combo1(0) = Detector
Combo1(1) = Yunit
Combo1(2) = Xunit
Combo1(3) = Xunit
Combo1(2) = Xunit
Text1(1) = Runtime
Text1(2) = TimeScale

'Ymax = 50
'Xmin = 0
'Ymin = 0
End If
End Sub

