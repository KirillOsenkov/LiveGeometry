VERSION 5.00
Begin VB.Form frmTips 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tip of the Day"
   ClientHeight    =   2985
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5730
   Icon            =   "frmTips.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrGradients 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3480
      Top             =   2400
   End
   Begin VB.CommandButton cmdPreviousTip 
      Caption         =   "&Previous Tip"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   765
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.Image imgLight 
         Height          =   480
         Left            =   75
         Picture         =   "frmTips.frx":030A
         Top             =   105
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblDidYouKnow 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   780
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   1755
         Left            =   90
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   3750
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      X1              =   280
      X2              =   376
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   280
      X2              =   376
      Y1              =   129
      Y2              =   129
   End
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const EnableTimer = 1

Dim Tips As New Collection

Const LB1 = 250
Const UB1 = 255
Const LB2 = 160
Const UB2 = 255

Dim C1 As Long, C2 As Long
Dim R1 As Long, R2 As Long
Dim G1 As Long, G2 As Long
Dim B1 As Long, B2 As Long
Dim R3 As Single, R4 As Single
Dim G3 As Single, G4 As Single
Dim B3 As Single, B4 As Single

Private Sub chkLoadTipsAtStartup_Click()
setShowTips = chkLoadTipsAtStartup.Value = 1
SaveSetting AppName, "General", "ShowTips", Format(-CInt(setShowTips))
End Sub

Private Sub cmdNextTip_Click()
DoNextTip
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdPreviousTip_Click()
DoNextTip True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdOK_Click
End Sub

Private Sub Form_Load()
FillDialogStrings
chkLoadTipsAtStartup.Value = -setShowTips
LoadTips
If setGradientFill Then GradientInit

Visible = True
cmdNextTip.SetFocus
End Sub

Public Sub DisplayCurrentTip()
UpdateBuffer
'If setCurrentTip >= 1 And setCurrentTip <= Tips.Count Then
'    lblTipText.Caption = Tips.Item(setCurrentTip)
'End If

'Gradient picContainer.hDC, vbWhite, RGB(128 + Rnd * 128, 128 + Rnd * 128, 128 + Rnd * 128), 0, 0, picContainer.ScaleWidth, picContainer.ScaleHeight, False
'picContainer.Refresh
End Sub

Private Sub DoNextTip(Optional ByVal Back As Boolean = False)
If Back Then
    setCurrentTip = setCurrentTip - 1
    If setCurrentTip < 1 Then setCurrentTip = Tips.Count
Else
    setCurrentTip = setCurrentTip + 1
    If Tips.Count < setCurrentTip Then setCurrentTip = 1
End If

DisplayCurrentTip
End Sub

'==================================

Private Sub UpdateBuffer()
Dim lpRect As RECT
Dim S As String
Const m = 8

If setGradientFill Then
    Gradient picContainer.hDC, C1, C2, 0, 0, picContainer.ScaleWidth, picContainer.ScaleHeight, False
Else
    picContainer.Cls
End If

If setCurrentTip >= 1 And setCurrentTip <= Tips.Count Then S = Tips(setCurrentTip) Else S = "No tip today!"
lpRect.Left = m
lpRect.Top = 3 * m + imgLight.Height
lpRect.Right = picContainer.ScaleWidth - 2 * m
lpRect.Bottom = picContainer.ScaleHeight - m
DrawText picContainer.hDC, S, Len(S), lpRect, DT_LEFT Or DT_TOP Or DT_WORDBREAK

S = GetString(ResTipDidYouKnow)
TextOut picContainer.hDC, 2 * m + imgLight.Width, imgLight.Height \ 2, S, Len(S)

picContainer.PaintPicture imgLight.Picture, m, m
End Sub

Private Sub GradientInit()
Const K As Single = 2
Randomize
R1 = LB1 + Rnd * (UB1 - LB1)
G1 = LB1 + Rnd * (UB1 - LB1)
B1 = LB1 + Rnd * (UB1 - LB1)
R2 = LB2 + Rnd * (UB2 - LB2)
G2 = LB2 + Rnd * (UB2 - LB2)
B2 = LB2 + Rnd * (UB2 - LB2)
R3 = (Rnd * 2 * K - K)
G3 = (Rnd * 2 * K - K)
B3 = (Rnd * 2 * K - K)
R4 = (Rnd * 2 * K - K)
G4 = (Rnd * 2 * K - K)
B4 = (Rnd * 2 * K - K)
C1 = RGB(R1, G1, B1)
C2 = RGB(R2, G2, B2)

UpdateBuffer

#If EnableTimer = 1 Then
    tmrGradients.Enabled = True
#End If

End Sub

Private Sub GradientChange()
R1 = R1 + R3
If R1 > UB1 Then R1 = UB1: R3 = -R3
If R1 < LB1 Then R1 = LB1: R3 = -R3
G1 = G1 + G3
If G1 > UB1 Then G1 = UB1: G3 = -G3
If G1 < LB1 Then G1 = LB1: G3 = -G3
B1 = B1 + B3
If B1 > UB1 Then B1 = UB1: B3 = -B3
If B1 < LB1 Then B1 = LB1: B3 = -B3

R2 = R2 + R4
If R2 > UB2 Then R2 = UB2: R4 = -R4
If R2 < LB2 Then R2 = LB2: R4 = -R4
G2 = G2 + G4
If G2 > UB2 Then G2 = UB2: G4 = -G4
If G2 < LB2 Then G2 = LB2: G4 = -G4
B2 = B2 + B4
If B2 > UB2 Then B2 = UB2: B4 = -B4
If B2 < LB2 Then B2 = LB2: B4 = -B4

C1 = RGB(R1, G1, B1)
C2 = RGB(R2, G2, B2)

UpdateBuffer
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting AppName, "General", "CurrentTip", setCurrentTip
End Sub

Private Sub tmrGradients_Timer()
GradientChange
End Sub

Public Sub FillDialogStrings()
cmdNextTip.Caption = GetString(ResTipNext)
cmdPreviousTip.Caption = GetString(ResTipPrevious)
chkLoadTipsAtStartup.Caption = GetString(ResShowTips)
Caption = GetString(ResTipOfTheDay)
End Sub
'====================================================

Public Sub LoadTips()
Tips.Add "Вы можете прокручивать рисунок клавишами влево, вправо, вверх и вниз на клавиатуре. Если при этом удерживать Ctrl, то прокрутка будет осуществляться быстрее."
Tips.Add "Вы можете увеличивать/уменьшать рисунок клавишами + и - на клавиатуре."
Tips.Add "Вы можете быстро включить или выключить сетку и оси координат, нажав F6 или F7 соответственно."
Tips.Add "Любые действия в DG, кроме операций с файлами, могут быть отменены. Количество отменяемых действий не ограничено. Ctrl+Z - отменить действие, Ctrl+R - вернуть отмененное действие."
Tips.Add "Щелчок правой кнопкой на любом из элементов рисунка показывает контекстное меню для этого элемента."
Tips.Add "F5 показывает калькулятор."
Tips.Add "Во время создания геометрических объектов (кроме многоугольника), щелчок правой кнопкой в пустом месте рисунка отменяет создание объекта. Повторный щелчок правой кнопкой делает активным инструмент Указатель."
Tips.Add "Ctrl+A - быстрое переключение ""Автопоказа имени точки"": отображать имя у вновь создаваемых точек или нет."
Tips.Add "Построить середину отрезка или измерить его длину можно простым щелчком по этому отрезку, когда активен соответствующий инструмент."
Tips.Add "Нажатие F9 вызывает диалог настроек DG."
Tips.Add "Панели инструментов, строку состояния и линейки можно скрывать в меню ""Вид"""
Tips.Add "Панель инструментов можно перемещать на любой край экрана за выпуклую область в ее левом / верхнем конце."
Tips.Add ""
Tips.Add ""
Tips.Add ""
Tips.Add ""
'Tips.Add ""


DoNextTip
End Sub

