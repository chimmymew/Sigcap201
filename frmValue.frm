VERSION 5.00
Begin VB.Form frmValue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "สัญญาณขาเข้า"
   ClientHeight    =   555
   ClientLeft      =   5400
   ClientTop       =   4965
   ClientWidth     =   3615
   Icon            =   "frmValue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   3615
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "ไม่พบการ์ด"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   -360
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
mainMDI.Timer1.Enabled = False
End Sub

