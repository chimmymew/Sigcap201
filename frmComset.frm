VERSION 5.00
Begin VB.Form frmComset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ตั้งค่าอุปกรณ์รับข้อมูล"
   ClientHeight    =   765
   ClientLeft      =   5400
   ClientTop       =   5160
   ClientWidth     =   3225
   Icon            =   "frmComset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   3225
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
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   735
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
      ItemData        =   "frmComset.frx":C75D
      Left            =   1320
      List            =   "frmComset.frx":C788
      TabIndex        =   3
      Text            =   "COM1"
      Top             =   360
      Width           =   975
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
      ItemData        =   "frmComset.frx":C7DC
      Left            =   1320
      List            =   "frmComset.frx":C7E6
      TabIndex        =   1
      Text            =   "Microsys Serial"
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ช่องเชื่อมต่อ"
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
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "การ์ดรับข้อมูล"
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
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmComset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change(Index As Integer)
Select Case Index
Case 0
    If Combo1(0).Text = "Microsys serial" Then
        Combo1(1).Enabled = True
        Else
        Combo1(1).Text = "USB"
        Combo1(1).Enabled = False
    End If
End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
Select Case Index
Case 0
    If Combo1(0).Text = "Microsys Serial" Then
        Combo1(1).Enabled = True
        Else
        Combo1(1).Text = "USB"
        Combo1(1).Enabled = False
    End If
End Select

End Sub

Private Sub Command1_Click()
On Error GoTo errdet:
    If Combo1(0).Text = "Microsys Serial" Then
    If mainMDI.mCom.PortOpen = True Then mainMDI.mCom.PortOpen = False
    mainMDI.mCom.CommPort = Val(Right(Combo1(1), 1))
    UseUSB = False
    If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True
    mainMDI.Timer1.Enabled = True
    Else
    UseUSB = True
    mainMDI.Timer1.Enabled = False
    If mainMDI.mCom.PortOpen = True Then mainMDI.mCom.PortOpen = False
    End If
    If UseUSB Then ConnectToHID mainMDI.hwnd
Unload Me
Exit Sub
errdet:
MsgBox Err.Description, vbInformation, "Error"
End Sub

Private Sub Form_Load()
mainMDI.Toolbar1.Buttons(6).Value = tbrPressed

End Sub




Private Sub Form_Unload(Cancel As Integer)
mainMDI.Toolbar1.Buttons(6).Value = tbrUnpressed
End Sub

