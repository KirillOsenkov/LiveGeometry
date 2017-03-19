VERSION 5.00
Begin VB.Form frmAboutSimple 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   6420
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5844
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.6
      Charset         =   204
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   61
   Icon            =   "frmAboutSimple.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   487
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrGradients 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5040
      Top             =   2640
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "More about the authors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   4215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://dg.osenkov.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   276
      Left            =   1680
      MouseIcon       =   "frmAboutSimple.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   2160
      Width           =   2328
   End
   Begin VB.Label lblSupport2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dg@osenkov.com"
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1920
      MouseIcon       =   "frmAboutSimple.frx":0594
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5280
      Width           =   1668
   End
   Begin VB.Label lblSupport 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help and technical support:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1320
      TabIndex        =   8
      Top             =   4920
      Width           =   3168
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 1999-2007 Sergey A. Rakov, Kirill Osenkov"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   4776
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coordinator:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   2160
      MouseIcon       =   "frmAboutSimple.frx":06E6
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3480
      Width           =   1056
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coordinator:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   3120
      Width           =   1056
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coordinator:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   2160
      MouseIcon       =   "frmAboutSimple.frx":0838
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2760
      Width           =   1056
   End
   Begin VB.Label lblDGSoftware 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dynamic Geometry Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   3630
   End
   Begin VB.Label lblDG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   1080
   End
   Begin VB.Label lblDG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Index           =   1
      Left            =   2880
      TabIndex        =   10
      Top             =   0
      Width           =   1080
   End
End
Attribute VB_Name = "frmAboutSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const EnableTimer = 0

Option Explicit

Const LB1 = 155
Const UB1 = 255
Const LB2 = 55
Const UB2 = 255

Dim C1 As Long, C2 As Long
Dim R1 As Long, R2 As Long
Dim G1 As Long, G2 As Long
Dim B1 As Long, B2 As Long
Dim R3 As Long, R4 As Long
Dim G3 As Long, G4 As Long
Dim B3 As Long, B4 As Long

Private Sub cmdMore_Click()
DisplayHelpTopic ResHlp_About
End Sub

Public Sub SayUserID()
MsgBox "User_ID=" & Trim(Str(Pervert(GetS))), vbInformation
Clipboard.Clear
Clipboard.SetText Trim(Str(Pervert(GetS)))
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then cmdOK_Click
End Sub

Private Sub Form_Load()
Dim Z As Long

Caption = GetString(ResAbout)
lblDGSoftware = GetString(ResTitles)
lblSupport = GetString(ResSupport)
lblVersion = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
cmdMore.Caption = GetString(ResTitles + 2 + 2 * 3)

lblTitle(0).ToolTipText = EMailRSA
lblTitle(2).ToolTipText = EMailOK

Center lblCopyright
Center lblVersion
Center lblDG(0)
Center lblDGSoftware
Center cmdOK
Center lblSupport
Center lblSupport2
Center cmdMore
Center lblURL
lblDG(1).Move lblDG(0).Left + 2, lblDG(0).Top + 2

For Z = 0 To 2
    lblTitle(Z).Caption = GetString(ResTitles + 2 * Z + 2)
    Center lblTitle(Z)
Next

If setGradientFill Then GradientInit
End Sub

Private Sub lblDG_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Shift = 7 Then SayUserID
End Sub

Private Sub lblSupport2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ShellExecute 0, "open", "mailto:" & lblSupport2.Caption, vbNullString, vbNullString, 1
Else
    '
End If
End Sub

Private Sub lblURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute 0, "open", lblURL.Caption, vbNullString, vbNullString, 1
End Sub

Private Sub lblTitle_Click(Index As Integer)
If Index = 0 Then
    ShellExecute 0, "open", "mailto:" & EMailRSA, vbNullString, vbNullString, 1
End If
If Index = 2 Then
    ShellExecute 0, "open", "mailto:" & EMailOK, vbNullString, vbNullString, 1
End If
End Sub

Private Sub Center(L1 As Object)
L1.Left = (ScaleWidth - L1.Width) / 2
End Sub

'=================================

Private Sub UpdateBuffer()
Gradient hDC, C1, C2, 0, 0, ScaleWidth, ScaleHeight, False
Refresh
End Sub

Private Sub GradientInit()
Const K = 2
Randomize
R1 = LB1 + Rnd * (UB1 - LB1)
G1 = LB1 + Rnd * (UB1 - LB1)
B1 = LB1 + Rnd * (UB1 - LB1)
R2 = LB2 + Rnd * (UB2 - LB2)
G2 = LB2 + Rnd * (UB2 - LB2)
B2 = LB2 + Rnd * (UB2 - LB2)
R3 = Int(Rnd * 2) * 2 * K - K
G3 = Int(Rnd * 2) * 2 * K - K
B3 = Int(Rnd * 2) * 2 * K - K
R4 = Int(Rnd * 2) * 2 * K - K
G4 = Int(Rnd * 2) * 2 * K - K
B4 = Int(Rnd * 2) * 2 * K - K
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

Private Sub tmrGradients_Timer()
GradientChange
End Sub
