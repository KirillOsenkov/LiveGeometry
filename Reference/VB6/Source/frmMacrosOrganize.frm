VERSION 5.00
Begin VB.Form frmMacrosOrganize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Упорядочить макросы"
   ClientHeight    =   4776
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5052
   HelpContextID   =   209
   Icon            =   "frmMacrosOrganize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4776
   ScaleWidth      =   5052
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Удалить макрос"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Добавить макрос"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.ListBox lstMacros 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2736
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   828
   End
   Begin VB.Label lblDesc 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   4815
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   4920
      Y1              =   4215
      Y2              =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   120
      X2              =   4920
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblList 
      Caption         =   "Список загруженных макросов:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmMacrosOrganize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim unlCancel As Boolean

Private Sub cmdAdd_Click()
If Not LoadMacroAs(True) Then Exit Sub
FillMacroList
lstMacros.ListIndex = lstMacros.ListCount - 1
End Sub

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdOK_Click()
unlCancel = False
Unload Me
End Sub

Private Sub cmdRemove_Click()
Dim tIndex As Long
tIndex = lstMacros.ListIndex

RemoveMacro tIndex + 1
FillMacroList

If lstMacros.ListCount > 0 Then
    If tIndex > lstMacros.ListCount - 1 Then tIndex = lstMacros.ListCount - 1
    lstMacros.ListIndex = tIndex
Else
    lblDesc.Caption = ""
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub Form_Load()
FillStrings
FillMacroList
If lstMacros.ListCount > 0 Then lstMacros.ListIndex = 0
End Sub

Public Sub FillStrings()
cmdCancel.Caption = GetString(ResCancel)
cmdAdd.Caption = GetString(ResMacroAdd)
cmdRemove.Caption = GetString(ResMacroRemove)
lblList.Caption = GetString(ResMacroList)
lblDescription.Caption = GetString(ResDescription)
Caption = GetString(ResMnuMacroOrganize)
End Sub

Public Sub FillMacroList()
Dim Z As Long

lstMacros.Clear
For Z = 1 To MacroCount
    lstMacros.AddItem Macros(Z).Name
Next

cmdRemove.Enabled = lstMacros.ListCount > 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub lstMacros_Click()
lblDesc.Caption = Macros(lstMacros.ListIndex + 1).Description
End Sub
