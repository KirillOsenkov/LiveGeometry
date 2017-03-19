VERSION 5.00
Begin VB.Form frmDemoProps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo properties"
   ClientHeight    =   5160
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5388
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   205
   Icon            =   "frmDemoProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClearDesc 
      Caption         =   "Clear"
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Top             =   3960
      Width           =   1812
   End
   Begin VB.HScrollBar hsbDelay 
      Height          =   255
      LargeChange     =   1000
      Left            =   2760
      Max             =   15000
      Min             =   500
      SmallChange     =   500
      TabIndex        =   5
      Top             =   840
      Value           =   5000
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   3372
   End
   Begin VB.ListBox lstObjects 
      Height          =   2640
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5280
      Y1              =   4455
      Y2              =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   5280
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblDelayValue 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.5 s"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3828
      TabIndex        =   10
      Top             =   1200
      Width           =   408
   End
   Begin VB.Label lblParticipate 
      Caption         =   "A checkmark near an item means it will take part in the demo as a stand-alone step."
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblMax 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "10 s"
      Height          =   195
      Left            =   4995
      TabIndex        =   8
      Top             =   1200
      Width           =   300
   End
   Begin VB.Label lblMin 
      AutoSize        =   -1  'True
      Caption         =   "0.5 s"
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label lblDelay 
      Caption         =   "Autodemo step delay:"
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image imgTimer 
      Height          =   384
      Left            =   2640
      Picture         =   "frmDemoProps.frx":0442
      Top             =   120
      Width           =   384
   End
   Begin VB.Label lblDescription 
      Caption         =   "Step description:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5175
   End
End
Attribute VB_Name = "frmDemoProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean
Dim DemoList As LinearObjectList

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdClearDesc_Click()
txtDescription.Text = GetObjectDescription(DemoList.Items(lstObjects.ListIndex + 1).Type, DemoList.Items(lstObjects.ListIndex + 1).Index)
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then unlCancel = True: Unload Me
End Sub

Private Sub Form_Load()
FillDialogStrings

DemoList = GenerateDemoSequence
FillListBoxWithLinearList lstObjects, DemoList
If nDemoInterval < 500 Or nDemoInterval > 15000 Then nDemoInterval = defDemoInterval
hsbDelay.Value = nDemoInterval

'ShadowControl cmdOK
'ShadowControl cmdCancel

FormMain.Enabled = False
unlCancel = False
Show
End Sub

Private Sub FillListBoxWithLinearList(lstList As ListBox, L As LinearObjectList)
Dim Z As Long
If L.Count = 0 Then Exit Sub

With lstList
    .Clear
    For Z = 1 To DemoList.Count
        .AddItem Z & ". " & GetObjectName(L.Items(Z).Type, L.Items(Z).Index)
        .Selected(Z - 1) = L.Items(Z).Participate
    Next
    .ListIndex = 0
End With

AddListboxScrollbar lstList
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Public Sub FillDialogStrings()
cmdCancel.Caption = GetString(ResCancel)
cmdClearDesc.Caption = GetString(ResDefault)
Caption = GetString(ResDemoOptions)
lblParticipate = GetString(ResDemoParticipatingItems)
lblDelay = GetString(ResDemoDelay)
lblMin = "0.5 " & GetString(ResSeconds)
lblDelayValue = "5 " & GetString(ResSeconds)
lblMax = "15 " & GetString(ResSeconds)
lblDescription = GetString(ResDemoStepDescription)
End Sub

Private Sub Form_Resize()
Line1.X2 = ScaleWidth - Line1.X1
Line2.X2 = Line1.X2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Z As Long
If unlCancel Then
    FormMain.Enabled = True
    FormMain.SetFocus
    Exit Sub
End If

For Z = 1 To DemoList.Count
    DemoList.Items(Z).Participate = lstObjects.Selected(Z - 1)
Next

UpdateObjectsWithLinearListData DemoList
nDemoInterval = hsbDelay.Value

FormMain.Enabled = True
FormMain.SetFocus

PaperCls
ShowAll
End Sub

Private Sub hsbDelay_Change()
lblDelayValue = Format(hsbDelay.Value / 1000, "#0.0") & " " & GetString(ResSeconds)
End Sub

Private Sub hsbDelay_Scroll()
hsbDelay_Change
End Sub

Private Sub lstObjects_Click()
txtDescription.Text = DemoList.Items(lstObjects.ListIndex + 1).Description
End Sub

Private Sub txtDescription_Change()
DemoList.Items(lstObjects.ListIndex + 1).Description = txtDescription.Text
End Sub
