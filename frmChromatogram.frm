VERSION 5.00
Begin VB.Form frmChromatogram 
   BackColor       =   &H8000000A&
   Caption         =   "กราฟและสัญญาณ"
   ClientHeight    =   4440
   ClientLeft      =   4215
   ClientTop       =   3975
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChromatogram.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   8670
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmChromatogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = "สัญญาณของ : " + Analysis_title
Form_Resize
mainMDI.Toolbar1.Buttons(7).Value = tbrPressed
End Sub

Private Sub Form_Resize()
graph_redraw

End Sub

Private Sub Picture1_Click()
graph_redraw

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Caption = Analysis_title + "  ตำแหน่ง (" + Format(x, "0.0") + "," + Format(y, "0.0000") + ")"

End Sub






Private Sub Form_Unload(Cancel As Integer)
mainMDI.Toolbar1.Buttons(7).Value = tbrUnpressed
End Sub

