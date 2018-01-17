VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmDataCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "การทำมาตรฐาน"
   ClientHeight    =   3105
   ClientLeft      =   5790
   ClientTop       =   4770
   ClientWidth     =   4740
   Icon            =   "FrmDataCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4740
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483647
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ที่"
         Object.Width           =   1481
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "ความเข้มข้น"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "หน่วย"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ค่าเฉลี่ย"
         Object.Width           =   1834
      EndProperty
   End
   Begin VB.TextBox Text2 
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
      Left            =   2280
      TabIndex        =   5
      Top             =   240
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
      ItemData        =   "FrmDataCal.frx":12FD4
      Left            =   1080
      List            =   "FrmDataCal.frx":12FED
      TabIndex        =   3
      Text            =   "mg/L"
      Top             =   240
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
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "วาดกราฟ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "เรียกมาตรฐานจากไฟล์"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "เพิ่ม"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "บันทึกมาตรฐาน"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "ค่าเฉลี่ยของสัญญาณ"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "เลือกหน่วย"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ความเข้มข้น"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "FrmDataCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iList As ListItem


Private Sub Label4_Click()

Dim i As Integer

On Error GoTo errdet:
cmd1.DialogTitle = "บันทึกมาตรฐาน"
cmd1.Filter = "Calibration File(*.calib)|*.calib"
cmd1.FileName = "calibra" + Format(filecount, "000")
cmd1.ShowSave

Open cmd1.FileName For Output As #1
For i = 1 To FrmDataCal.ListView1.ListItems.Count
With FrmDataCal.ListView1.ListItems(i)
Print #1, .Text + "," + .SubItems(1) + "," + .SubItems(2) + "," + .SubItems(3)
End With
Next


Close #1
Exit Sub
errdet:
Close #1
End Sub

Private Sub Label5_Click()
If Text1 <> "" And Text2 <> "" Then

Set iList = FrmDataCal.ListView1.ListItems.Add(, , "#" + Trim(Str(FrmDataCal.ListView1.ListItems.Count + 1)))
With iList
.SubItems(1) = Str(Val(Text1))
.SubItems(2) = Combo1
.SubItems(3) = Str(Val(Text2))
End With

End If
Redraw_Calibration
End Sub

Private Sub Label6_Click()
On Error GoTo errdet:
cmd1.DialogTitle = "เปิดมาตรฐานที่เก็บไว้"
cmd1.Filter = "Calibration file (*.calib)|*.calib"
cmd1.CancelError = True
cmd1.ShowOpen
Dim dat_in As String
Dim fld() As String
FrmDataCal.ListView1.ListItems.Clear

Open cmd1.FileName For Input As #1

While Not EOF(1)
Line Input #1, dat_in
fld = Split(dat_in, ",")
Set iList = FrmDataCal.ListView1.ListItems.Add(, , fld(0))
With iList
.SubItems(1) = fld(1)
.SubItems(2) = fld(2)
.SubItems(3) = fld(3)
End With

Wend



Close #1

Redraw_Calibration

Exit Sub
errdet:
Close #1
End Sub

Private Sub Label7_Click()
Redraw_Calibration
End Sub

Private Sub Listview1_DblClick()
Dim i As Integer
On Error GoTo errdet:
FrmDataCal.ListView1.ListItems.Remove FrmDataCal.ListView1.SelectedItem.Index

For i = 1 To FrmDataCal.ListView1.ListItems.Count
 FrmDataCal.ListView1.ListItems(i).Text = "#" + Trim(Str(i))
Next
errdet:
End Sub



