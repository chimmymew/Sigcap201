VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5280
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ไม่พบการ์ดรับสัญญาณ!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
End Sub
