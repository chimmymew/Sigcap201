VERSION 5.00
Begin VB.Form frmInjection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Ѻ�ѭ�ҳ���"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "frmInjection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4020
   Begin VB.CommandButton Command1 
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "��ҹ����� 0 �ҷ�"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmInjection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case Xunit
Case "�Թҷ�"
    Timer1.Interval = 1000 * TimeScale
Case "�ҷ�"
    Timer1.Interval = 60000 * TimeScale
End Select

If Command1.Caption = "�����" Then
Timer1.Enabled = True
Command1.Caption = "��ش"

Else
Timer1.Enabled = False
frmPump.Timer2.Enabled = False
If frmPump.Command1.Caption = "��ش" Then frmPump.Command1.Value = True
Command1.Caption = "�����"
End If
frmValue.Show
End Sub

Private Sub Timer1_Timer()
getVal
Ti = Ti + 1
ReDim Preserve time_plot(Ti)
ReDim Preserve data_plot(Ti)
time_plot(Ti) = Ti * TimeScale
data_plot(Ti) = Current_Value
Label1.Caption = Format(Current_Value, "0.000") + " " + Yunit
'If Current_Value < Ymin Then Ymin = Ymin - 10
'If Current_Value > Ymax Then Ymax = Ymax + 10
graph_redraw
Label2.Caption = "��ҹ����� " + Format(Ti * TimeScale, "0.00") + " " + Xunit
If Ti * TimeScale > Xmax Then
Command1_Click
MsgBox "�ú��˹���������!", vbInformation, "˹�ҵ�ҧ�ʴ���ͤ���"
End If
End Sub
