VERSION 5.00
Begin VB.Form frmGraphControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "การควบคุมกราฟ"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "frmGraphControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   3555
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
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   0
      Width           =   615
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
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
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
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ค่าสูงที่สุดของแกนตั้ง"
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
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ค่าต่ำที่สุดของแกนตั้ง"
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
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmGraphControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Ymin = Text1(0)
Ymax = Text1(1)
graph_redraw
End Sub

Private Sub Form_Load()
Text1(0) = Ymin
Text1(1) = Ymax
mainMDI.Toolbar1.Buttons(8).Value = tbrPressed
End Sub


Private Sub Form_Unload(Cancel As Integer)
mainMDI.Toolbar1.Buttons(8).Value = tbrUnpressed
End Sub
