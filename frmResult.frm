VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResult 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ผลการวิเคราะห์สัญญาณ"
   ClientHeight    =   3465
   ClientLeft      =   5400
   ClientTop       =   4965
   ClientWidth     =   5025
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5025
   Begin VB.CommandButton Command3 
      Caption         =   "เทียบมาตรฐาน"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "กราฟมาตรฐาน"
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
      Left            =   840
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ที่"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "ตำแหน่ง"
         Object.Width           =   1481
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "เริ่ม"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "สิ้นสุด"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "พื้นที่"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "ความสูง"
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total_signal As Single
Dim Total_count As Integer
Dim Average As Single
Dim sum_Square As Single
Dim stdDeviation As Single
Dim delta_sumSquare() As Single
Dim percentageDeviation As Single
Dim Assume As Single

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
FrmDataCal.Show
frmCalibration.Show
frmCalibration.Height = 5000
frmCalibration.Width = 5000
End Sub

Private Sub Command3_Click()
If RegA <> 0 Then
Assume = (Average - RegB) / RegA
frmVariant.Text1.Text = frmVariant.Text1.Text + vbCrLf
frmVariant.Text1.Text = frmVariant.Text1.Text + "ค่าเฉลี่ยเทียบมาตรฐานได้ : " + Format(Assume, "0.00") + " " + cXunit + vbCrLf

If Assume > cXmax Then
frmVariant.Text1.Text = frmVariant.Text1.Text + "ระวัง ! ค่าที่คำนวณได้ตกนอกกราฟมาตรฐาน" + vbCrLf
frmVariant.Text1.ForeColor = vbRed
End If

Redraw_Calibration
frmCalibration.Picture1.DrawWidth = 2
frmCalibration.Picture1.ForeColor = vbGreen
frmCalibration.Picture1.Circle (Assume, Average), cXbound / 8
frmCalibration.Picture1.DrawWidth = 1
frmCalibration.Picture1.DrawStyle = DrawStyleConstants.vbDot
frmCalibration.Picture1.Line (0, Average)-(Assume, Average), vbGreen
frmCalibration.Picture1.Line (Assume, 0)-(Assume, Average), vbGreen
frmCalibration.Picture1.DrawStyle = DrawStyleConstants.vbSolid
End If
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Total_signal = 0
'delta_sumSquare(0) = 0
Total_count = 0
If frmAnalysis.Option1.Value = True Then
frmVariant.Text1.Text = ""
frmVariant.Text1.Text = frmVariant.Text1.Text + "ที่เลือก   ความสูง " + vbCrLf
frmVariant.Text1.Text = frmVariant.Text1.Text + "------------------------------" + vbCrLf
For i = 1 To ListView1.ListItems.Count
If ListView1.ListItems(i).Checked = True Then
frmVariant.Text1.Text = frmVariant.Text1.Text + ListView1.ListItems(i).Text + vbTab + ListView1.ListItems(i).SubItems(5) + vbCrLf
Total_signal = Total_signal + Val(ListView1.ListItems(i).SubItems(5))
Total_count = Total_count + 1
ReDim Preserve delta_sumSquare(Total_count)
delta_sumSquare(Total_count) = Val(ListView1.ListItems(i).SubItems(5))
End If
Next
If Total_count <> 0 Then
Average = Total_signal / Total_count
frmVariant.Text1.Text = frmVariant.Text1.Text + "------------------------------" + vbCrLf
frmVariant.Text1.Text = frmVariant.Text1.Text + "จำนวนสัญญาณทั้งสิ้น : " + vbTab + Trim(Str(Total_count)) + " สัญญาณ" + vbCrLf
frmVariant.Text1.Text = frmVariant.Text1.Text + "ค่าเฉลี่ยของทุกสัญญาณ : " + vbTab + Format(Average, "0.0000") + vbCrLf
delta_sumSquare(0) = 0
For i = 1 To Total_count
 delta_sumSquare(0) = delta_sumSquare(0) + (Average - delta_sumSquare(i)) ^ 2
Next

stdDeviation = Sqr(delta_sumSquare(0) / Total_count)
percentageDeviation = 100 * stdDeviation / Average
frmVariant.Text1.Text = frmVariant.Text1.Text + "ส่วนเบี่ยงเบนมาตรฐาน : " + vbTab + Format(stdDeviation, "0.0000") + vbCrLf
frmVariant.Text1.Text = frmVariant.Text1.Text + "ความแปรปรวนของสัญญาณ : " + vbTab + Format(stdDeviation ^ 2, "0.0000") + vbCrLf

Select Case percentageDeviation
        Case Is < 5
        frmVariant.Text1.Text = frmVariant.Text1.Text + "สัญญาณที่เลือกมามีความน่าเชื่อถือมาก (เบี่ยงเบน " + Format(percentageDeviation, "0.00") + "%)" + vbCrLf
                 frmVariant.Text1.ForeColor = vbBlue

        Case 5 To 10
         frmVariant.Text1.Text = frmVariant.Text1.Text + "สัญญาณที่เลือกมามีความน่าเชื่อถือปานกลาง (เบี่ยงเบน " + Format(percentageDeviation, "0.00") + "%)" + vbCrLf
                  frmVariant.Text1.ForeColor = vbBlack
      
        Case Is > 10
         frmVariant.Text1.Text = frmVariant.Text1.Text + "ระวังสัญญาณที่เลือกมาไม่มีความน่าเชื่อถือเลย (เบี่ยงเบน " + Format(percentageDeviation, "0.00") + "%)" + vbCrLf
         frmVariant.Text1.ForeColor = vbRed
End Select
End If

Else

frmVariant.Text1.Text = ""
frmVariant.Text1.Text = frmVariant.Text1.Text + "ที่เลือก      พื้นที่ " + vbCrLf
frmVariant.Text1.Text = frmVariant.Text1.Text + "------------------------------" + vbCrLf
For i = 1 To ListView1.ListItems.Count
If ListView1.ListItems(i).Checked = True Then
frmVariant.Text1.Text = frmVariant.Text1.Text + ListView1.ListItems(i).Text + vbTab + ListView1.ListItems(i).SubItems(4) + vbCrLf
Total_signal = Total_signal + Val(ListView1.ListItems(i).SubItems(4))
Total_count = Total_count + 1
ReDim Preserve delta_sumSquare(Total_count)
delta_sumSquare(Total_count) = Val(ListView1.ListItems(i).SubItems(4))
End If
Next

If Total_count <> 0 Then

Average = Total_signal / Total_count
frmVariant.Text1.Text = frmVariant.Text1.Text + "------------------------------" + vbCrLf
frmVariant.Text1.Text = frmVariant.Text1.Text + "จำนวนสัญญาณทั้งสิ้น : " + vbTab + Trim(Str(Total_count)) + " สัญญาณ" + vbCrLf
frmVariant.Text1.Text = frmVariant.Text1.Text + "ค่าเฉลี่ยของทุกสัญญาณ : " + vbTab + Format(Average, "0.0000") + vbCrLf
delta_sumSquare(0) = 0
For i = 1 To Total_count
 delta_sumSquare(0) = delta_sumSquare(0) + (Average - delta_sumSquare(i)) ^ 2
Next

stdDeviation = Sqr(delta_sumSquare(0) / Total_count)
percentageDeviation = 100 * stdDeviation / Average
frmVariant.Text1.Text = frmVariant.Text1.Text + "ส่วนเบี่ยงเบนมาตรฐาน : " + vbTab + Format(stdDeviation, "0.0000") + vbCrLf
frmVariant.Text1.Text = frmVariant.Text1.Text + "ความแปรปรวนของสัญญาณ : " + vbTab + Format(stdDeviation ^ 2, "0.0000") + vbCrLf

Select Case percentageDeviation
        Case Is < 5
        frmVariant.Text1.Text = frmVariant.Text1.Text + "สัญญาณที่เลือกมามีความน่าเชื่อถือมาก (เบี่ยงเบน " + Format(percentageDeviation, "0.00") + "%)" + vbCrLf
                 frmVariant.Text1.ForeColor = vbBlue

        Case 5 To 10
         frmVariant.Text1.Text = frmVariant.Text1.Text + "สัญญาณที่เลือกมามีความน่าเชื่อถือปานกลาง (เบี่ยงเบน " + Format(percentageDeviation, "0.00") + "%)" + vbCrLf
                  frmVariant.Text1.ForeColor = vbBlack
      
        Case Is > 10
         frmVariant.Text1.Text = frmVariant.Text1.Text + "ระวังสัญญาณที่เลือกมาไม่มีความน่าเชื่อถือเลย (เบี่ยงเบน " + Format(percentageDeviation, "0.00") + "%)" + vbCrLf
         frmVariant.Text1.ForeColor = vbRed
End Select
End If
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
graph_redraw
For i = peak_start(Val(Item.Index)) + 1 To peak_end(Val(Item.Index))
frmChromatogram.Picture1.Line (time_plot(i), data_plot(i))-(time_plot(i), Baseline), vbRed
Next
RotateText frmChromatogram.Picture1, "#" + Str(Val(Item.Index)) + "<-" + Str(peak_pos(Val(Item.Index))), peak_pos(Val(Item.Index)), peak_height(Val(Item.Index)), 90


End Sub
