VERSION 5.00
Begin VB.Form frmFileProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File properties"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   HelpContextID   =   206
   Icon            =   "frmFileProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFileName 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmFileProps.frx":030A
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame fraDescription 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmFileProps.frx":031F
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmFileProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdStatistics_Click()
Dim S As String
S = S & GetString(ResNumberOfItems) & vbCrLf
S = S & "_____________________" & vbCrLf
S = S & vbCrLf
S = S & GetString(ResToolPoints) & ": " & PointCount & vbCrLf
S = S & GetString(ResFigures) & ": " & VisualFigureCount & vbCrLf
S = S & GetString(ResLabels) & ": " & LabelCount & vbCrLf
S = S & GetString(ResButtons) & ": " & ButtonCount
MsgBox S, vbInformation + vbOKOnly, cmdStatistics.Caption
End Sub

Private Sub cmdStatistics_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Dir(DrawingName) <> "" Then
    ShellExecute 0, "open" & vbNullChar, "notepad.exe" & vbNullChar, DrawingName & vbNullChar, vbNullString, SW_MAXIMIZE
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then unlCancel = True: Unload Me
End Sub

Private Sub Form_Load()
FormMain.Enabled = False
unlCancel = False

Caption = GetString(ResFileProps)
cmdCancel.Caption = GetString(ResCancel)
fraDescription.Caption = GetString(ResComment)
cmdStatistics.Caption = GetString(ResStatistics)
txtFileName.Text = GetString(ResName) & ": " & DrawingName
If Dir(DrawingName) <> "" Then
    txtFileName = txtFileName.Text & "    " & GetString(ResSize) & ": " & Format(FileLen(DrawingName), "# ### ##0 ") & GetString(ResBytes)
End If

txtDescription.Text = nDescription

'ShadowControl cmdCancel
'ShadowControl cmdOK
'ShadowControl cmdStatistics

Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormMain.Enabled = True
FormMain.SetFocus
If unlCancel Then Exit Sub

nDescription = txtDescription.Text

PaperCls
ShowAll
Paper.Refresh
End Sub
