VERSION 5.00
Begin VB.UserControl ctlCalculator 
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   7710
   ToolboxBitmap   =   "ctlCalculator.ctx":0000
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Rnd"
      Height          =   375
      Index           =   35
      Left            =   3600
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Random"
      Height          =   375
      Index           =   36
      Left            =   3600
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Round"
      Height          =   375
      Index           =   33
      Left            =   2760
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Log"
      Height          =   375
      Index           =   26
      Left            =   2760
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Exp"
      Height          =   375
      Index           =   27
      Left            =   2760
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "/"
      Height          =   375
      Index           =   54
      Left            =   1200
      TabIndex        =   17
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "^"
      Height          =   375
      Index           =   24
      Left            =   840
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "."
      Height          =   375
      Index           =   49
      Left            =   480
      TabIndex        =   12
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "0"
      Height          =   375
      Index           =   48
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "*"
      Height          =   375
      Index           =   53
      Left            =   1200
      TabIndex        =   16
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "3"
      Height          =   375
      Index           =   47
      Left            =   840
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "2"
      Height          =   375
      Index           =   46
      Left            =   480
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "1"
      Height          =   375
      Index           =   45
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "-"
      Height          =   375
      Index           =   52
      Left            =   1200
      TabIndex        =   15
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "6"
      Height          =   375
      Index           =   44
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "5"
      Height          =   375
      Index           =   43
      Left            =   480
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      Height          =   375
      Index           =   42
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "+"
      Height          =   375
      Index           =   51
      Left            =   1200
      TabIndex        =   14
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "9"
      Height          =   375
      Index           =   41
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "8"
      Height          =   375
      Index           =   40
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "7"
      Height          =   375
      Index           =   39
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "e"
      Height          =   375
      Index           =   23
      Left            =   4560
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "PI"
      Height          =   375
      Index           =   22
      Left            =   4560
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00FFC0C0&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Ctg"
      Height          =   375
      Index           =   32
      Left            =   1920
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Tg"
      Height          =   375
      Index           =   31
      Left            =   1920
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Cos"
      Height          =   375
      Index           =   30
      Left            =   1920
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Sin"
      Height          =   375
      Index           =   29
      Left            =   1920
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcButton 
      BackColor       =   &H00C0C0FF&
      Caption         =   "[ ]"
      Height          =   375
      Index           =   38
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Rad"
      Height          =   375
      Index           =   37
      Left            =   4560
      TabIndex        =   34
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Min"
      Height          =   375
      Index           =   20
      Left            =   3600
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Abs"
      Height          =   375
      Index           =   28
      Left            =   2760
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Fact"
      Height          =   375
      Index           =   34
      Left            =   7560
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkInv 
      Caption         =   "Inv"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Sqr"
      Height          =   375
      Index           =   25
      Left            =   2760
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "°"
      Height          =   375
      Index           =   21
      Left            =   4560
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Max"
      Height          =   375
      Index           =   19
      Left            =   3600
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "If"
      Height          =   375
      Index           =   18
      Left            =   3600
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   6360
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Area"
      Height          =   375
      Index           =   16
      Left            =   5400
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   7320
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Arg"
      Height          =   375
      Index           =   14
      Left            =   6720
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   7320
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Norm"
      Height          =   375
      Index           =   12
      Left            =   6720
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   7320
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Y"
      Height          =   375
      Index           =   10
      Left            =   6720
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6360
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6360
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6360
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   7320
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "X"
      Height          =   375
      Index           =   8
      Left            =   6720
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "XAngle"
      Height          =   375
      Index           =   6
      Left            =   5400
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "OAngle"
      Height          =   375
      Index           =   4
      Left            =   5400
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Angle"
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdCalcButton 
      Caption         =   "Dist"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   5160
      X2              =   5160
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   5175
      X2              =   5175
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   1695
      X2              =   1695
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   1680
      X2              =   1680
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      Visible         =   0   'False
      X1              =   0
      X2              =   7080
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   1
      Visible         =   0   'False
      X1              =   0
      X2              =   7080
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   7080
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   7080
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Label lblHint 
      Caption         =   "Input [ to insert a dynamic measurement into the label text."
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   2040
      Width           =   7215
   End
End
Attribute VB_Name = "ctlCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_Textbox As TextBox
Dim m_IsInitialized As Boolean
Dim m_ShowTooltips As Boolean
Dim m_BracketsVisible As Boolean

Public TempFunctionName As String
Public TempParamCount As Long

Public Event ObjectChoiceComplete()

Public Property Get ParentTextbox() As TextBox
If m_IsInitialized Then Set ParentTextbox = m_Textbox
End Property

Public Property Set ParentTextbox(vTextBox As TextBox)
Set m_Textbox = vTextBox
m_IsInitialized = True
End Property

Public Function GetSelection() As String
If m_IsInitialized Then
    GetSelection = m_Textbox.SelText
End If
End Function

Public Sub SetSelection(ByVal S As String)
If m_IsInitialized Then
    m_Textbox.SelText = S
End If
End Sub

Public Sub InsertIntoTextbox(ByVal sStr As String, Optional ByVal InsertPoint As Long = -1)
Dim A As String, Z As Long

If m_IsInitialized Then
    With m_Textbox
        If .SelLength <> 0 Then
            Z = .SelStart
            A = .SelText
            A = Left(sStr, Len(sStr) + InsertPoint) & A & Right(sStr, -InsertPoint)
            .SelText = A
            .SelStart = Z
            .SelLength = Len(A)
        Else
            .SelLength = 0
            .SelText = sStr
            .SelStart = .SelStart + InsertPoint
        End If
        
        .SetFocus
    End With
End If
End Sub

Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vEnabled As Boolean)
Dim Z As Long
UserControl.Enabled = vEnabled
For Z = 0 To cmdCalcButton.UBound
    cmdCalcButton(Z).Enabled = vEnabled
Next
End Property

Private Sub chkInv_Click()
If chkInv.Value = 1 Then
    cmdCalcButton(29).Caption = "Arcsin"
    cmdCalcButton(30).Caption = "Arccos"
    cmdCalcButton(31).Caption = "Arctg"
    cmdCalcButton(32).Caption = "Arcctg"
Else
    cmdCalcButton(29).Caption = "Sin"
    cmdCalcButton(30).Caption = "Cos"
    cmdCalcButton(31).Caption = "Tg"
    cmdCalcButton(32).Caption = "Ctg"
End If
End Sub

Private Sub cmdCalcButton_Click(Index As Integer)
Dim A As String, Z As Long
Z = -1
Select Case Index
Case 0
    A = "Dist(,)"
    Z = InStr(A, ",") - Len(A) - 1
Case 1
    GoAndFetchMeSomePoints "Dist", 2
    Exit Sub
Case 3
    GoAndFetchMeSomePoints "Angle", 3
    Exit Sub
Case 5
    GoAndFetchMeSomePoints "OAngle", 3
    Exit Sub
Case 7
    GoAndFetchMeSomePoints "XAngle", 2
    Exit Sub
Case 9
    GoAndFetchMeSomePoints ".X", 1
    Exit Sub
Case 11
    GoAndFetchMeSomePoints ".Y", 1
    Exit Sub
Case 13
    GoAndFetchMeSomePoints "Norm", 1
    Exit Sub
Case 15
    GoAndFetchMeSomePoints "Arg", 1
    Exit Sub
Case 17
    GoAndFetchMeSomePoints "Area", 0
    Exit Sub
Case 2
    A = "Angle(,,)"
    Z = InStr(A, ",") - Len(A) - 1
Case 4
    A = "OAngle(,,)"
    Z = InStr(A, ",") - Len(A) - 1
Case 6
    A = "XAngle(,)"
    Z = InStr(A, ",") - Len(A) - 1
Case 8, 10
    A = "." & cmdCalcButton(Index).Caption
    Z = 0
Case 16, 18
    A = cmdCalcButton(Index).Caption & "(,,)"
    Z = InStr(A, ",") - Len(A) - 1
Case 18
    A = "If(,,)"
    Z = InStr(A, ",") - Len(A) - 1
Case 12, 14, 25, 27 To 32, 34, 37
    A = cmdCalcButton(Index).Caption & "()"
Case 19, 20, 26, 36
    A = cmdCalcButton(Index).Caption & "(,)"
    Z = InStr(A, ",") - Len(A) - 1
Case 21
    A = DegreeSign
    Z = 0
    If GetSelection <> "" Then
        A = "()" & A
        Z = -2
    End If
Case 22 To 24, 35, 39 To 54
    SetSelection ""
    A = cmdCalcButton(Index).Caption
    Z = 0
Case 33
    A = "Round(," & setNumberPrecision & ")"
    Z = InStr(A, ",") - Len(A) - 1
Case 38
    A = "[]"
Case Else
    Exit Sub
End Select

InsertIntoTextbox A, Z
End Sub

Public Property Get ShowTooltips() As Boolean
ShowTooltips = m_ShowTooltips
End Property

Public Property Let ShowTooltips(ByVal vShow As Boolean)
m_ShowTooltips = vShow
LoadTooltips vShow
End Property

Private Sub LoadTooltips(Optional ByVal ShouldClear As Boolean = False)
Dim Z As Long

If ShouldClear Then
    For Z = cmdCalcButton.LBound To cmdCalcButton.UBound
        cmdCalcButton(Z).ToolTipText = ""
    Next
Else
    
End If
End Sub

Public Sub ObjectSelectionComplete_CalcPoints()
Dim S As String, Z As Long

If Left(TempFunctionName, 1) <> "." Then
    S = TempFunctionName & "("
    For Z = 1 To TempObjectSelection.PointCount
        S = S & BasePoint(TempObjectSelection.Points(Z)).Name & IIf(Z = TempObjectSelection.PointCount, ")", ",")
    Next
Else
    S = BasePoint(TempObjectSelection.Points(1)).Name & TempFunctionName
End If

InsertIntoTextbox S, 0
TempFunctionName = ""
TempParamCount = 0
RaiseEvent ObjectChoiceComplete
Set GlobalCalc = Nothing
End Sub

Public Sub ObjectSelectionCancel_CalcPoints()
TempFunctionName = ""
TempParamCount = 0
m_Textbox.SetFocus
Set GlobalCalc = Nothing
End Sub

Public Sub GoAndFetchMeSomePoints(Optional ByVal FName As String = "Dist", Optional ByVal PCount As Long = 2)
Dim TempOSC As ObjectSelectionCaller
TempFunctionName = FName
TempParamCount = PCount
ObjectListClear TempObjectSelection
TempObjectSelection.PointCountMax = PCount
Select Case UserControl.Parent.Name
Case "frmLabelProps"
    TempOSC = oscCalcLabels
Case "frmAnPoint"
    TempOSC = oscCalcAnPoint
Case "frmCalculator"
    TempOSC = oscCalculator
End Select
Set GlobalCalc = Me
ObjectSelectionBegin ostCalcPoints, False, TempOSC
End Sub

Public Property Get BracketsVisible() As Boolean
BracketsVisible = m_BracketsVisible
End Property

Public Property Let BracketsVisible(ByVal vNewValue As Boolean)
m_BracketsVisible = vNewValue
cmdCalcButton(38).Visible = vNewValue
ShowHint IIf(vNewValue, GetString(ResCalcBase + 2 * ResCalcBrackets), "")
End Property

Public Sub ShowHint(ByVal S As String)
If lblHint.Caption <> S Then lblHint.Caption = S
End Sub

Private Sub cmdCalcButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
i = Index
If chkInv.Value And i >= ResCalcSin And i <= ResCalcCtg Then
    i = i + ResCalcASin - ResCalcSin
    ShowHint GetString(ResCalcBase + 2 * i)
    Exit Sub
End If

If i > 38 Then
    ShowHint cmdCalcButton(i).Caption
Else
    ShowHint GetString(ResCalcBase + 2 * i)
End If
End Sub

Private Sub cmdHelp_Click()
DisplayHelpTopic ResHlp_Interface_CalcPanel
End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowHint GetString(ResHelp)
End Sub

Private Sub UserControl_Initialize()
m_BracketsVisible = True
ShowHint GetString(ResCalcBase + 2 * ResCalcBrackets)
If PointCount < 3 Then
    cmdCalcButton(ResCalcAngle).Enabled = False
    cmdCalcButton(ResCalcAngleB).Enabled = False
    cmdCalcButton(ResCalcOAngle).Enabled = False
    cmdCalcButton(ResCalcOAngleB).Enabled = False
    If PointCount < 2 Then
        cmdCalcButton(ResCalcDistance).Enabled = False
        cmdCalcButton(ResCalcDistanceB).Enabled = False
        cmdCalcButton(ResCalcXangle).Enabled = False
        cmdCalcButton(ResCalcXangleB).Enabled = False
        cmdCalcButton(ResCalcArea).Enabled = False
        cmdCalcButton(ResCalcAreaB).Enabled = False
        If PointCount = 0 Then
            cmdCalcButton(ResCalcX).Enabled = False
            cmdCalcButton(ResCalcXB).Enabled = False
            cmdCalcButton(ResCalcY).Enabled = False
            cmdCalcButton(ResCalcYB).Enabled = False
            cmdCalcButton(ResCalcNorm).Enabled = False
            cmdCalcButton(ResCalcNormB).Enabled = False
            cmdCalcButton(ResCalcArg).Enabled = False
            cmdCalcButton(ResCalcArgB).Enabled = False
        End If
    End If
End If
End Sub

Private Sub UserControl_InitProperties()
BracketsVisible = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowHint ""
End Sub

Private Sub UserControl_Paint()
Dim qrc As RECT
'UserControl.Cls
qrc.Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
qrc.Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
DrawEdge UserControl.hDC, qrc, EDGE_ETCHED, BF_RECT
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
BracketsVisible = PropBag.ReadProperty("BracketsVisible", False)
End Sub

Private Sub UserControl_Resize()
Dim TPP As Long

TPP = Screen.TwipsPerPixelX
Line1(0).X2 = UserControl.ScaleWidth - Line1(0).X1
Line1(1).X2 = Line1(0).X2
Line2(0).X2 = Line1(0).X2
Line2(1).X2 = Line1(0).X2

Line1(0).Y1 = UserControl.ScaleHeight - 3 * TPP ' cmdCalcButton (32).Top + cmdCalcButton(32).Height + cmdCalcButton(38).Top
Line1(0).Y2 = Line1(0).Y1
Line2(0).Y1 = Line1(0).Y1 + 1 * TPP
Line2(0).Y2 = Line2(0).Y1

lblHint.Move cmdHelp.Left, ScaleHeight - lblHint.Height - 8 * TPP, ScaleWidth - 2 * cmdHelp.Left
End Sub

Public Sub InsertBrackets()
cmdCalcButton_Click ResCalcBrackets
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "BracketsVisible", BracketsVisible, False
End Sub
