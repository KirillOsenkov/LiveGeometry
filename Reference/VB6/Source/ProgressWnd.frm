VERSION 5.00
Begin VB.Form frmProgressWnd 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Progress"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   236
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmProgressWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Offs = 10
Const ForeGradient = vbWhite
Dim ProgressColor1 As Long, ProgressColor2 As Long

Private Sub Form_Load()
On Local Error Resume Next
Dim TW As Long, TH As Long
'BringToTop hWnd
Randomize
If setGradientFill Then
    ProgressColor1 = vbInactiveTitleBar
    ProgressColor2 = vbWhite
Else
    ProgressColor1 = RGB(192, 192, 192)
End If
TW = TextWidth(ProgressMsg)
TH = TextHeight(ProgressMsg)
Width = (TW + Offs * 4 + 4) * Screen.TwipsPerPixelX
Height = (2 * TH + Offs * 5 + 8) * Screen.TwipsPerPixelY
lblCaption.Caption = ProgressMsg
lblCaption.Move (ScaleWidth - TW) \ 2, 2 * Offs + 2, TW, TH
picProgress.Move lblCaption.Left, lblCaption.Top + lblCaption.Height + Offs, TW, TH + 4
Visible = True
Refresh
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
Dim lpRect As RECT
If setGradientFill Then Gradient hDC, ForeGradient, BackColor, 0, 0, ScaleWidth, ScaleHeight, False Else Cls
lpRect.Right = ScaleWidth
lpRect.Bottom = ScaleHeight
DrawEdge hDC, lpRect, EDGE_RAISED, BF_RECT
lpRect.Right = ScaleWidth - Offs
lpRect.Bottom = ScaleHeight - Offs
lpRect.Left = Offs
lpRect.Top = Offs
DrawEdge hDC, lpRect, EDGE_ETCHED, BF_RECT
End Sub

Public Sub Progress(ByVal Percentage As Double)
On Local Error Resume Next
If Percentage > 1 Then Percentage = 1
Dim hBrush As Long, lpRect As RECT, tStr As String, SW As Long, SH As Long, LenTStrX As Long, LenTStrY As Long
SW = picProgress.ScaleWidth
SH = picProgress.ScaleHeight

hBrush = CreateSolidBrush(picProgress.BackColor)
lpRect.Right = SW
lpRect.Bottom = SH
lpRect.Left = SW * Percentage
FillRect picProgress.hDC, lpRect, hBrush
DeleteObject hBrush

If setGradientFill Then
    Gradient picProgress.hDC, ProgressColor1, ProgressColor2, 0, 0, SW * Percentage, SH
Else
    picProgress.Line (0, 0)-(SW * Percentage, SH), ProgressColor1, BF
End If
tStr = Int(Percentage * 100) & "%"
LenTStrX = picProgress.TextWidth(tStr)
LenTStrY = picProgress.TextHeight(tStr)

picProgress.ForeColor = 0
TextOut picProgress.hDC, (SW - LenTStrX) / 2, (SH - LenTStrY) / 2, tStr, Len(tStr)
picProgress.Refresh
End Sub
