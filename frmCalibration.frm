VERSION 5.00
Begin VB.Form frmCalibration 
   BackColor       =   &H00FFFFFF&
   Caption         =   "กราฟมาตรฐาน"
   ClientHeight    =   2655
   ClientLeft      =   5805
   ClientTop       =   4980
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmCalibration.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   3855
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmCalibration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Picture1.Width = Me.Width - 130
Picture1.Height = Me.Height - 400
Picture1.Cls
Redraw_Calibration
End Sub
