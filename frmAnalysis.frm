VERSION 5.00
Begin VB.Form frmAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "พารามิเตอร์การวิเคราะห์สัญญาณ"
   ClientHeight    =   1110
   ClientLeft      =   4995
   ClientTop       =   5160
   ClientWidth     =   4740
   Icon            =   "frmAnalysis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4740
   Begin VB.OptionButton Option2 
      Caption         =   "ตามพื้นที่"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ตามความสูง"
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
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "วิเคราะห์"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   360
      Width           =   975
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
      Left            =   2760
      TabIndex        =   3
      Text            =   "5"
      Top             =   360
      Width           =   855
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
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Text            =   "0.5"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ค้นหา"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "พื้นสัญญาณต่ำสุดที่นับ (treshold)"
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
      Left            =   -240
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "สัญญาณพื้นฐาน (base line)"
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
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo errdet:
Baseline = Format(findMin(1, UBound(data_plot)) + (findMax(1, UBound(data_plot)) / 10), "0.0000")
Treshold = Format(findMax(1, UBound(data_plot)) / 10, "0.0000")
Text1(0) = Str(Baseline)
Text1(1) = Str(Treshold)
errdet:
End Sub

Private Sub Command2_Click()
Baseline = Val(Text1(0))
Treshold = Val(Text1(1))
count_peak
process_peak
End Sub

Private Sub Form_Load()
mainMDI.Toolbar1.Buttons(4).Value = tbrPressed
Text1(0) = Str(Baseline)
Text1(1) = Str(Treshold)

End Sub



Private Sub Form_Unload(Cancel As Integer)
mainMDI.Toolbar1.Buttons(4).Value = tbrUnpressed

End Sub
