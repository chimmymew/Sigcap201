VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRGB 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "แหล่งแสง RGB"
   ClientHeight    =   1395
   ClientLeft      =   1380
   ClientTop       =   1770
   ClientWidth     =   5835
   Icon            =   "frmRGB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5835
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "400"
      Top             =   960
      Width           =   375
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   -120
      TabIndex        =   2
      ToolTipText     =   "เลือกแสง"
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Min             =   400
      Max             =   700
      SelStart        =   400
      TickStyle       =   3
      Value           =   400
      TextPosition    =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ปรับศูนย์"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "เรนเดอร์"
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
      Left            =   3600
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Wavelenght:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "nm."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lbl_rgb 
      BackColor       =   &H80000004&
      Caption         =   "RGB=[127,0,255]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "frmRGB.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim R, G, B, Wavelength As Integer

Private Sub Command1_Click()
Dim Mystring As String

If mainMDI.mCom.PortOpen = False Then mainMDI.mCom.PortOpen = True
mainMDI.Timer1.Enabled = False
Command1.Caption = "รอ"
Command1.Enabled = False
Do

DoEvents
Loop While CommReady = False

mainMDI.mCom.Output = "R"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0
'Debug.Print mainMDI.mCom.OutBufferCount
Mystring = Trim(Str((R)))
mainMDI.mCom.Output = Mystring
mainMDI.mCom.Output = Chr(13)
Do
DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0


mainMDI.mCom.Output = "G"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0
Mystring = Trim(Str(G))
mainMDI.mCom.Output = Mystring
mainMDI.mCom.Output = Chr(13)
Do
DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0

mainMDI.mCom.Output = "B"

Do

DoEvents
Loop While InStr(1, Comm_buffer, ">") = 0

Mystring = Trim(Str(B))
mainMDI.mCom.Output = Mystring
mainMDI.mCom.Output = Chr(13)
Do
DoEvents
Loop While InStr(1, Comm_buffer, "*") = 0

mainMDI.Timer1.Enabled = True
Command1.Caption = "เรนเดอร์"
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Zero_Cal = Current_Value
End Sub

Private Sub Slider1_Change()

Wavelength = Slider1.Value
Text1 = Slider1
    
RGB_out
End Sub

Private Sub Text1_Change()
 Wavelength = Val(Text1)
If Wavelength < 400 Then Wavelength = 400
If Wavelength > 700 Then Wavelength = 700
Slider1.Value = Wavelength
RGB_out
End Sub

Private Sub RGB_out()

Dim Mystring As String

Select Case Wavelength
    Case 400 To 549
    R = 0
    G = (Wavelength - 400)
    B = 150 - G
    
    Case 550 To 700
    B = 0
    R = (Wavelength - 550)
    G = 150 - R
    End Select
    lbl_rgb.Caption = "RGB=[" + Str(Int(R)) + "," + Str(Int(G)) + "," + Str(Int(B)) + "]"
    'Picture1.BackColor = RGB(R + 100, G + 100, B + 100)
    'mCom.Output = "w"
    'mCom.Output = Chr(R + 32) + Chr(B + 32) + Chr(G + 32) + vbCr
End Sub
