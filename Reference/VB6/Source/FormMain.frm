VERSION 5.00
Begin VB.Form FormMain 
   BackColor       =   &H8000000C&
   Caption         =   "DG"
   ClientHeight    =   4836
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   5796
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFF00&
   HelpContextID   =   1
   Icon            =   "FormMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin DG.ctlMenuBar MenuBar 
      Align           =   1  'Align Top
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   5790
      _ExtentX        =   10224
      _ExtentY        =   868
      BackColor       =   12632256
      ForeColor       =   16776960
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   41
      ScaleWidth      =   483
      GradientColor   =   192
      GradientInverse =   -1  'True
      WhatsThisHelp   =   -1  'True
   End
   Begin VB.Timer tmrDemo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   3960
   End
   Begin DG.ctlMenuBar MenuBar 
      Align           =   1  'Align Top
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   495
      Width           =   5790
      _ExtentX        =   10224
      _ExtentY        =   868
      BackColor       =   12632256
      ForeColor       =   16776960
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   42
      ScaleWidth      =   483
      GradientColor   =   16508571
      GradientInverse =   -1  'True
      WhatsThisHelp   =   -1  'True
   End
   Begin VB.PictureBox Docked 
      Align           =   4  'Align Right
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   3285
      MousePointer    =   9  'Size W E
      ScaleHeight     =   296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   984
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.PictureBox RulerButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   480
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox XRuler 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   840
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   4245
      Begin VB.Line IndicX 
         DrawMode        =   6  'Mask Pen Not
         X1              =   37
         X2              =   37
         Y1              =   4
         Y2              =   8
      End
   End
   Begin VB.PictureBox YRuler 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3270
      Left            =   240
      ScaleHeight     =   271
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   165
      Begin VB.Line IndicY 
         DrawMode        =   6  'Mask Pen Not
         X1              =   1
         X2              =   9
         Y1              =   69
         Y2              =   69
      End
   End
   Begin VB.Timer tmrClock 
      Interval        =   445
      Left            =   4440
      Top             =   3780
   End
   Begin VB.PictureBox StatusBar 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   479
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4536
      Width           =   5790
      Begin VB.Label Status 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X=0.0; Y=0.0"
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
         Index           =   2
         Left            =   960
         TabIndex        =   4
         Top             =   15
         Width           =   1140
      End
      Begin VB.Label Clock 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12:00:00"
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
         Left            =   4305
         TabIndex        =   3
         Top             =   30
         Width           =   750
      End
      Begin VB.Label Status 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   60
      End
   End
   Begin VB.PictureBox Canvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2340
      Index           =   0
      Left            =   600
      MouseIcon       =   "FormMain.frx":0442
      MousePointer    =   99  'Custom
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1950
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "S&ave as"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Begin VB.Menu mnuBitmap 
            Caption         =   "&Bitmap"
         End
         Begin VB.Menu mnuMetafile 
            Caption         =   "&Metafile"
         End
         Begin VB.Menu mnuEnhancedMetafile 
            Caption         =   "&Enhanced metafile"
         End
         Begin VB.Menu mnuExportSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuJSPHTML 
            Caption         =   "&Java Sketchpad HTML"
         End
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMRUFile 
         Caption         =   "MRUFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "&Clear all"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertLabel 
         Caption         =   "&Insert text label"
      End
      Begin VB.Menu mnuInsertButton 
         Caption         =   "Insert &button"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Калькулятор"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProps 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuFullscreen 
         Caption         =   "Fullscreen"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWE 
         Caption         =   "&Show watch expressions window"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowMainbar 
         Caption         =   "Show main toolbar"
      End
      Begin VB.Menu mnuShowToolbar 
         Caption         =   "Show toolbar"
      End
      Begin VB.Menu mnuShowStatusbar 
         Caption         =   "Show status bar"
      End
      Begin VB.Menu mnuShowRulers 
         Caption         =   "&Rulers"
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowGrid 
         Caption         =   "&Grid"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuShowAxes 
         Caption         =   "Axes"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuViewSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDemoOptions 
         Caption         =   "Demo options"
      End
      Begin VB.Menu mnuDemo 
         Caption         =   "Step-by-step &demo"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuFigures 
      Caption         =   "Figures"
      Begin VB.Menu mnuFigPoints 
         Caption         =   "Point"
         Begin VB.Menu mnuToolPoint 
            Caption         =   "Point"
         End
         Begin VB.Menu mnuToolPointOnFigure 
            Caption         =   "Figure point"
         End
      End
      Begin VB.Menu mnuFigLines 
         Caption         =   "Lines"
         Begin VB.Menu mnuToolSegment 
            Caption         =   "Segment"
         End
         Begin VB.Menu mnuToolRay 
            Caption         =   "Ray"
         End
         Begin VB.Menu mnuToolLine 
            Caption         =   "Line"
         End
         Begin VB.Menu mnuToolParallelLine 
            Caption         =   "Parallel line"
         End
         Begin VB.Menu mnuToolPerpendicularLine 
            Caption         =   "Perpendicular line"
         End
         Begin VB.Menu mnuToolBisector 
            Caption         =   "Bisector"
         End
      End
      Begin VB.Menu mnuFigCircles 
         Caption         =   "Circles"
         Begin VB.Menu mnuToolCircle 
            Caption         =   "Circle"
         End
         Begin VB.Menu mnuToolCircleByRadius 
            Caption         =   "Circle by radius"
         End
         Begin VB.Menu mnuToolArc 
            Caption         =   "Arc"
         End
      End
      Begin VB.Menu mnuFigConstruction 
         Caption         =   "Construction"
         Begin VB.Menu mnuToolMiddlePoint 
            Caption         =   "Middle point"
         End
         Begin VB.Menu mnuToolSymmPoint 
            Caption         =   "Symmetric point"
         End
         Begin VB.Menu mnuToolReflectedPoint 
            Caption         =   "Reflected point"
         End
         Begin VB.Menu mnuToolInvert 
            Caption         =   "Inverted point"
         End
         Begin VB.Menu mnuToolIntersect 
            Caption         =   "Intersection points"
         End
         Begin VB.Menu mnuConstructionSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPolygon 
            Caption         =   "Polygon"
         End
         Begin VB.Menu mnuDynamicLocus 
            Caption         =   "Dynamic Locus"
         End
      End
      Begin VB.Menu mnuFigSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnalytic 
         Caption         =   "Analytic"
         Begin VB.Menu mnuAnPoint 
            Caption         =   "Point"
         End
         Begin VB.Menu mnuAnLine 
            Caption         =   "Line"
         End
         Begin VB.Menu mnuAnCircle 
            Caption         =   "Circle"
         End
         Begin VB.Menu mnuFigSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVector 
            Caption         =   "Vector"
         End
         Begin VB.Menu mnuBezier 
            Caption         =   "Bezier"
         End
         Begin VB.Menu mnuAnSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuActiveAxes 
            Caption         =   "Add active axes"
         End
      End
      Begin VB.Menu mnuFigMeasure 
         Caption         =   "Measurement"
         Begin VB.Menu mnuToolMeasureDistance 
            Caption         =   "Distance"
         End
         Begin VB.Menu mnuToolMeasureAngle 
            Caption         =   "Angle"
         End
         Begin VB.Menu mnuToolMeasureArea 
            Caption         =   "Area"
         End
      End
      Begin VB.Menu mnuFigSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPointList 
         Caption         =   "Point list"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFigureList 
         Caption         =   "Figure list"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuMacros 
      Caption         =   "Macros"
      Begin VB.Menu mnuMacroCreate 
         Caption         =   "Create a macro"
      End
      Begin VB.Menu mnuMacroResults 
         Caption         =   "Select results"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMacroSave 
         Caption         =   "Save macro"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMacroSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMacroLoad 
         Caption         =   "Load macro"
      End
      Begin VB.Menu mnuMacroOrganize 
         Caption         =   "Упорядочить макросы"
      End
      Begin VB.Menu mnuMacroSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMacroRun 
         Caption         =   "Macro Name"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuOptionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLanguage 
         Caption         =   "&Language"
         Begin VB.Menu mnuLangEnglish 
            Caption         =   "&English"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuLangUkrainian 
            Caption         =   "&Ukrainian"
         End
         Begin VB.Menu mnuLangRussian 
            Caption         =   "&Russian"
         End
         Begin VB.Menu mnuLangGerman 
            Caption         =   "&German"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Help contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuTip 
         Caption         =   "Tip of the day"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "Debug"
      Visible         =   0   'False
      Begin VB.Menu mnuUndoBufferSize 
         Caption         =   "Show undo buffer size"
      End
   End
   Begin VB.Menu mnuFigurePopup 
      Caption         =   "Figure popup"
      Visible         =   0   'False
      Begin VB.Menu mnuChooseFigure 
         Caption         =   "Choose figure"
         Visible         =   0   'False
         Begin VB.Menu mnuFigureChoice 
            Caption         =   "Fig1"
            Index           =   1
         End
      End
      Begin VB.Menu mnuVectorProperties 
         Caption         =   "Vector properties"
         Visible         =   0   'False
         Begin VB.Menu mnuVectorProp 
            Caption         =   "Vector1"
            Index           =   1
         End
      End
      Begin VB.Menu mnuVectorDelete 
         Caption         =   "Delete vector"
         Visible         =   0   'False
         Begin VB.Menu mnuVectorDel 
            Caption         =   ""
            Index           =   1
         End
      End
      Begin VB.Menu mnuFigureProperties 
         Caption         =   "Figure properties"
      End
      Begin VB.Menu mnuFigureSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHideFigure 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuMeasurementProperties 
         Caption         =   "Measurement properties"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDeleteFigure 
         Caption         =   "Delete figure"
      End
      Begin VB.Menu mnuDeleteMeasurement 
         Caption         =   "Delete measurement"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPointPopup 
      Caption         =   "Point popup"
      Visible         =   0   'False
      Begin VB.Menu mnuChoosePoint 
         Caption         =   "Choose point"
         Visible         =   0   'False
         Begin VB.Menu mnuPointChoice 
            Caption         =   "Point1"
            Index           =   1
         End
      End
      Begin VB.Menu mnuPointProperties 
         Caption         =   "Point properties"
      End
      Begin VB.Menu mnuShowPointName 
         Caption         =   "Show name"
      End
      Begin VB.Menu mnuSnapToFigure 
         Caption         =   "Snap to figure"
         Begin VB.Menu mnuSnapTo 
            Caption         =   "Snap fig 1"
            Index           =   1
         End
      End
      Begin VB.Menu mnuReleasePoint 
         Caption         =   "Release point"
      End
      Begin VB.Menu mnuLocusProps 
         Caption         =   "Locus"
         Begin VB.Menu mnuCreateLocus 
            Caption         =   "Create locus"
         End
         Begin VB.Menu mnuClearLocus 
            Caption         =   "Clear locus"
         End
      End
      Begin VB.Menu mnuPointSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHidePoint 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuDeletePoint 
         Caption         =   "Delete point"
      End
   End
   Begin VB.Menu mnuLabelPopup 
      Caption         =   "Label popup"
      Visible         =   0   'False
      Begin VB.Menu mnuLabelProperties 
         Caption         =   "Label properties"
      End
      Begin VB.Menu mnuRecalcLabel 
         Caption         =   "Recalculate"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLabelSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFixLabel 
         Caption         =   "Fix in place"
      End
      Begin VB.Menu mnuDeleteLabel 
         Caption         =   "Delete label"
      End
   End
   Begin VB.Menu mnuChooseMacroObject 
      Caption         =   "Choose macro object"
      Visible         =   0   'False
      Begin VB.Menu mnuMacroObject 
         Caption         =   "MacroObject"
         Index           =   1
      End
   End
   Begin VB.Menu mnuSGPopup 
      Caption         =   "SGPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuSGProperties 
         Caption         =   "SG properties"
      End
      Begin VB.Menu mnuSGSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSGDelete 
         Caption         =   "Delete SG"
      End
   End
   Begin VB.Menu mnuButtonPopup 
      Caption         =   "ButtonPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuButtonProperties 
         Caption         =   "Button properties"
      End
      Begin VB.Menu mnuButtonSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButtonMovable 
         Caption         =   "Movable"
      End
      Begin VB.Menu mnuButtonDelete 
         Caption         =   "Delete button"
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Status2PositionRatio = 1 / 2
Const MinDocked = 80

Dim Redocking As Boolean
'Dim HoverItem As Long
Dim OX As Long
Dim OldDockedWidth As Long
Dim IsFullScreen As Boolean
Dim StatusBarResizeEngaged As Boolean
Public FormResizeEngaged As Boolean
Dim MenubarResizeEngaged As Boolean

Public StatusBarSpecialMode As Boolean
Public PaperCursor As CursorState

'==================================================
'           Map of canvas to Paper
'==================================================

Private Sub Canvas_DblClick(Index As Integer)
PaperDoubleClick
End Sub

Private Sub Canvas_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PaperMouseDown Button, Shift, X, Y
End Sub

Private Sub Canvas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PaperMouseMove Button, Shift, X, Y
End Sub

Private Sub Canvas_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PaperMouseUp Button, Shift, X, Y
End Sub

Private Sub Canvas_Paint(Index As Integer)
PaperPaint
End Sub

Private Sub Canvas_Resize(Index As Integer)
PaperResize
End Sub

'==================================================

Private Sub Clock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then ShowStatus StrConv(Format(Now, "Long Date"), vbProperCase)
End Sub

Private Sub mnuCalculator_Click()
MenuCommand ResCalcBase
End Sub

Private Sub mnuFigures_Click()
If DragS.State = dscNormalState Then
    mnuFigureList.Enabled = VisualFigureCount > 0
    mnuPointList.Enabled = PointCount > 0
End If
End Sub

Private Sub mnuMacroOrganize_Click()
MenuCommand ResMnuMacroOrganize
End Sub

Private Sub mnuMeasurementProperties_Click()
MenuCommand ResMnuMeasurementProperties
End Sub

Private Sub tmrClock_Timer()
Static OldTime As String
If Time <> OldTime Then
    OldTime = Time
    If setShowClock Then Clock.Caption = Time
End If
End Sub

'==================================================
'                               DOCKED
'==================================================

'Private Sub Docked_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If X < 2 And Button = 1 Then
'    Redocking = True
'    DrawStripe X
'    OX = X
'End If
'End Sub
'
'Private Sub Docked_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim P As POINTAPI
'If Redocking Then
'    If Docked.Width - X > ScaleWidth / 2 Or Docked.Width - X < MinDocked Then Exit Sub
'    DrawStripe OX
'    DrawStripe X
'    OX = X
'End If
'End Sub
'
'Private Sub Docked_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Redocking Then
'    Redocking = False
'    DrawStripe OX
'    If Docked.Width - X > ScaleWidth / 2 Then
'        X = Docked.Width - ScaleWidth / 2
'    End If
'    If Docked.Width - X < MinDocked Then
'        X = Docked.Width - MinDocked
'    End If
'    Docked.Width = Docked.Width - X
'    Form_Resize
'    'ValueTable1.Resize
'End If
'End Sub
'
'Public Sub Docked_Resize()
'If Docked.ScaleWidth < 4 Then Exit Sub
''ValueTable1.Move 2, 0, Docked.ScaleWidth - 2, Docked.ScaleHeight
'End Sub
'
'==================================================
'   Form keydown procedures
'==================================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim F As String, Z As Long
Dim X As Double, Y As Double, nX As Double, nY As Double, Index As Long, T As Double
Dim tempTree As Tree, lpPoint1 As POINTAPI
Dim DontLine As Boolean, CurTool As Long

If IsFullScreen Then
    Me.Fullscreen = False
    Exit Sub
End If

On Error GoTo EH:
Select Case KeyCode
Case vbKeyA To vbKeyE, vbKeyG To vbKeyZ
    PaperKeyDown KeyCode, Shift
    
    If KeyCode = vbKeyO And Shift = 7 Then
        If InputBox("Enter function:", "f(x)") = Format(2 ^ 32, "#0") Then ShowAbout
        Exit Sub
    End If
    
    If KeyCode = vbKeyZ And (Shift And 2) = 2 And Not mnuUndo.Enabled Then
        If DragS.State = dscMacroStateGivens Or DragS.State = dscMacroStateResults Then i_CancelMacro: Exit Sub
        If DragS.State = dscMacroStateRun Then CancelMacroRun: Exit Sub
        CancelOperation
        DragS.ShouldSkipUndo = True
    End If
    
    If KeyCode = vbKeyR And DragS.State = dscNormalState Then
        PaperCls
        ShowAll
    End If
    
    If KeyCode = vbKeyX And Shift = 0 And DragS.State = dscNormalState Then mnuAnPoint_Click
    
Case vbKeyEscape
    If DragS.State = dscNormalState Then
        PaperCls
        ShowAll
        EnableMenus mnsStandard
        If DrawingState = dsSelect Then Exit Sub
        SelectTool dsSelect
        'If DrawingState > dsMeasureAngle Then DrawingState = dsSelect Else Menus(2).Items(DrawingState + 2).Checked = False
        'DrawingState = dsSelect
        'Menus(2).Items(DrawingState + 2).Checked = True
        'MenuBar(2).Refresh
        SetMousePointer curStateArrow
        ShowStatus GetString(ResSelect)
    Else
        CancelOperation
    End If
    
'Case vbKeyF5
'    MenuCommand ResCalcBase
'

Case vbKeyF
    If Shift <> 0 Or DragS.State <> dscNormalState Then Exit Sub
'    If Shift And vbCtrlMask Then
        F = UCase(InputBox("f(x) = "))
        If F = "" Then Exit Sub
        ShowStatus GetString(ResWorkingPleaseWait)

        'On Error Resume Next
        nX = CanvasBorders.P1.X
        Index = AddCustomVariable("X", nX)
        tempTree = BuildTree(F)
        If IsSevere(WasThereAnErrorEvaluatingLastExpression) Then
            MsgBox GetString(ResError) & ": " & vbCrLf & GetString(ResEvalErrorBase + WasThereAnErrorEvaluatingLastExpression * 2 - 2) & ".", vbExclamation
            WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
            DeleteCustomVariable "X"
            ShowStatus
            Exit Sub
        End If
        nY = tempTree.Branches(1).CurrentValue
        ToPhysical nX, nY
        If Abs(nY) < 1000 Then
            SetPixelV Paper.hDC, nX, nY, colPlotColor
            MoveToEx Paper.hDC, nX, nY, lpPoint1
            DontLine = True
        Else
            DontLine = False
        End If
        On Error GoTo EH

        Dim DC As Long
        'T = timeGetTime
        DC = Paper.hDC
        Paper.ForeColor = colPlotColor

        For X = CanvasBorders.P1.X To CanvasBorders.P2.X Step (CanvasBorders.P2.X - CanvasBorders.P1.X) / PaperScaleWidth
            CustomVariables(Index).CurrentValue = X
            Y = RecalculateTree(tempTree, 1)
            If WasThereAnErrorEvaluatingLastExpression Then
                DontLine = True
                WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
                GoTo NextX
            End If
            nX = X
            nY = Y
            ToPhysical nX, nY
            If Abs(nY) < 1000 And Not DontLine Then LineTo DC, nX, nY
            If DontLine And Abs(nY) < 1000 Then MoveToEx DC, nX, nY, lpPoint1 'SetPixelV DC, NX, NY, colPlotColor
            DontLine = False
            If Abs(nY) > 1000 Then DontLine = True
NextX: Next

        'MsgBox timeGetTime - T
        Paper.Refresh

        DeleteCustomVariable "X"
        ShowStatus
    
'Case vbKeyD
'    If Shift And vbCtrlMask Then
'        f = UCase(InputBox("f(x) = "))
'        MsgBox "(" & f & ")' = " & Differentiate(f)
'    End If
    
Case Else
    PaperKeyDown KeyCode, Shift
    
End Select
Exit Sub

EH:
ERR.Clear
If Index = 0 Then Resume Next
If X < CanvasBorders.P2.X Then
    X = X + 0.1
    CustomVariables(Index).CurrentValue = X
    DontLine = True
    Resume
Else
    Resume Next
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
PaperKeyUp KeyCode, Shift
End Sub

'================================
' Application entry point.
'================================

Public Sub Form_Load()

MenubarResizeEngaged = True

PreparePaths
ParseCommandLine
FillFonts

Set Paper = Canvas(0)
CreateTempFontDC

GetSettings
GetDefaults
Me.PrepareFormControls
CanSave = IsOKVersion

FillMenuTransposition
CreateMenus
FillStrings
FillMRU
InitGraphics

modDrawing.InitDrawing

ProcessAutoloadMacros

'========================================================================

LockWindowUpdate hWnd
Me.Visible = True

LoadIdentity
ScrollCanvas
PaperCls
ShowAll

If CommandLineFile <> "" Then OpenCommandLineFile
LockWindowUpdate 0

'========================================================================

MenubarResizeEngaged = False

'========================================================================

Load frmSettings

#If conTips = 1 Then
    If setShowTips Then frmTips.Show
#End If

#If conProtected = 1 Then
    MsgBox GetString(ResBuyDG), vbInformation, "www.dg.osenkov.com"
#End If

'========================================================================
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowStatus
HideCooWhenMouseMoving
End Sub

Public Sub HideCooWhenMouseMoving()
HideIndicX
HideIndicY
Status(2).Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If IsDirty Then
    Select Case MsgBox(GetString(ResSave) & " " & DrawingName & "?", vbYesNoCancel + vbQuestion, GetString(ResMsgConfirmation))
        Case vbYes
            MenuCommand ResSave
        Case vbNo
            'do nothing
        Case vbCancel
            Cancel = 1
            Exit Sub
    End Select
End If
End Sub

Public Sub Form_Resize()
On Local Error Resume Next
Const DrawShadows As Boolean = False
Dim ptRect As RECT, ptRect2 As RECT, Z As Long, MenubarResizeWasEngaged As Boolean
Dim TheMargin As Long

If FormResizeEngaged Or WindowState = vbMinimized Then Exit Sub
FormResizeEngaged = True
MenubarResizeWasEngaged = MenubarResizeEngaged
MenubarResizeEngaged = False

If Me.Fullscreen Then TheMargin = 0 Else TheMargin = Margin

ptRect.Left = TheMargin
ptRect.Top = TheMargin
If setShowRulers Then
    ptRect.Left = ptRect.Left + YRuler.Width
    ptRect.Top = ptRect.Top + XRuler.Height
End If
ptRect.Right = ScaleWidth - TheMargin
ptRect.Bottom = ScaleHeight - TheMargin
If IsFullScreen Then
    GetWindowRect Me.hWnd, ptRect2
    ptRect.Left = ptRect.Left - ptRect2.Left
    ptRect.Top = ptRect.Top - ptRect2.Top
    ptRect.Right = GetSystemMetrics(SM_CXSCREEN) + ptRect.Left
    ptRect.Bottom = GetSystemMetrics(SM_CYSCREEN) + ptRect.Top
End If
If setShowStatusbar <> StatusBar.Visible Then StatusBar.Visible = setShowStatusbar
If StatusBar.Visible Then ptRect.Bottom = ptRect.Bottom - StatusBar.Height

For Z = MenuBar.LBound To MenuBar.UBound
    If Z = 1 And MenuBar(1).Visible <> setShowMainbar Then MenuBar(1).Visible = setShowMainbar
    If Z = 2 And MenuBar(2).Visible <> setShowToolbar Then MenuBar(2).Visible = setShowToolbar
    If MenuBar(Z).Visible Then
        Select Case MenuBar(Z).Align
            Case vbAlignTop
                ptRect.Top = ptRect.Top + MenuBar(Z).Height
            Case vbAlignBottom
                ptRect.Bottom = ptRect.Bottom - MenuBar(Z).Height
            Case vbAlignLeft
                ptRect.Left = ptRect.Left + MenuBar(Z).Width
            Case vbAlignRight
                ptRect.Right = ptRect.Right - MenuBar(Z).Width
            Case Else
        End Select
    End If
Next
If Docked.Visible Then
    ptRect.Right = ptRect.Right - Docked.Width
    'ValueTable1.Resize
End If
If ptRect.Left >= ptRect.Right Then
    ptRect.Right = ptRect.Left + TheMargin
End If
If ptRect.Top >= ptRect.Bottom Then
    ptRect.Bottom = ptRect.Top + TheMargin
End If

Paper.Move ptRect.Left, ptRect.Top, ptRect.Right - ptRect.Left, ptRect.Bottom - ptRect.Top
XRuler.Move Paper.Left, Paper.Top - XRuler.Height - 1, Paper.Width
YRuler.Move Paper.Left - YRuler.Width - 1, Paper.Top, YRuler.Width, Paper.Height
RulerButton.Move YRuler.Left, XRuler.Top, YRuler.Width, XRuler.Height

If DrawShadows Then
    Cls
    If setGradientFill Then
        If RulerButton.Visible Then
            ptRect2.Left = YRuler.Left
            ptRect2.Top = YRuler.Top - XRuler.Height - 1
            ptRect2.Right = YRuler.Width + ptRect2.Left
            ptRect2.Bottom = XRuler.Height + ptRect2.Top
            ShadowRect hDC, ptRect2, BackColor
        End If
        If XRuler.Visible Then
            ptRect2.Left = Paper.Left
            ptRect2.Top = Paper.Top - XRuler.Height - 1
            ptRect2.Right = Paper.Width + ptRect2.Left
            ptRect2.Bottom = XRuler.Height + ptRect2.Top
            ShadowRect hDC, ptRect2, BackColor
        End If
        If YRuler.Visible Then
            ptRect2.Left = Paper.Left - YRuler.Width - 1
            ptRect2.Top = Paper.Top
            ptRect2.Right = YRuler.Width + ptRect2.Left
            ptRect2.Bottom = Paper.Height + ptRect2.Top
            ShadowRect hDC, ptRect2, BackColor
        End If
        ShadowRect hDC, ptRect, BackColor
    End If
End If

IndicX.Y2 = XRuler.Height
IndicY.X2 = YRuler.Width

MenubarResizeEngaged = MenubarResizeWasEngaged
FormResizeEngaged = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteTempFontDC
SaveSettings
SaveLastPaths
CD.HelpCommand = HELP_QUIT
CD.ShowHelp
End
End Sub

'===================================================
' Process menu item clicks
'===================================================

Private Sub MenuBar_Command(Index As Integer, ByVal MenuNumber As Integer, ByVal MenuText As String)
If ToolTipShown Then Unload frmToolTip: DoEvents

If Index = 1 Then
    MenuCommand MenuBar(1).ItemAuxIndex(MenuNumber)
End If

If Index = 2 Then
    Select Case MenuBar(Index).ToolbarMode
    
    '==============================================
    '                               TOOLBAR
    '==============================================
    Case mbsToolBar
        SelectTool Transpose(MenuNumber, MenuTransposition)
'        If Transpose(MenuNumber, MenuTransposition) = dsPolygon Then
'            MenuCommand ResStaticObjectBase + 2 * sgPolygon
'        End If
    
    '==============================================
    '                               SELECT OBJECTS FINISH
    '==============================================
    Case mbsSelectObjectsFinish
        If MenuNumber = 1 Then
            ObjectSelectionComplete
        ElseIf MenuNumber = 2 Then
            ObjectSelectionCancel
        End If
    
    '==============================================
    '                               CANCEL
    '==============================================
    Case mbsCancel
        ObjectSelectionCancel
    
    '==============================================
    '                               DEMO
    '==============================================
    Case mbsDemo
        Select Case MenuNumber
        Case 1 'First frame
            DemoFirstStep
        Case 2 'Previous frame
            DemoPreviousStep
        Case 3 'Next frame
            DemoNextStep
        Case 4 'Last frame
            DemoLastStep
        Case 5 'Autoplay
            AutorunDemo
            MenuBar(2).CheckItem 5, Not MenuBar(2).ItemChecked(5)
            MenuBar(2).Refresh
        Case 6 'Continue
            EndDemo
        End Select
    
    '==============================================
    '                               MACRO GIVENS
    '==============================================
    Case mbsMacroGivens
        If MenuNumber = 1 Then
            MacroCreateResultsInit
        ElseIf MenuNumber = 2 Then
            i_CancelMacro
        End If
        
    '==============================================
    '                               MACRO RESULTS
    '==============================================
    Case mbsMacroResults
        If MenuNumber = 1 Then
            MacroSaveInit
        ElseIf MenuNumber = 2 Then
            i_CancelMacro
        End If
        
    '==============================================
    '                               MACRO RUN
    '==============================================
    Case mbsMacroRun
        CancelMacroRun
        
    End Select
End If
End Sub

Public Sub SelectTool(ByVal ToolNumber As DrawState)
MenuBar(2).CheckItem TransposeInv(DrawingState, MenuTransposition), False
DrawingState = ToolNumber
MenuBar(2).CheckItem TransposeInv(ToolNumber, MenuTransposition), True
MenuBar(2).Refresh
ImitateMouseMove
End Sub

Private Sub MenuBar_CommandHover(Index As Integer, ByVal MenuNumber As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Shift As Integer)
Select Case MenuBar(Index).ToolbarMode
Case mbsToolBar
    If MenuNumber > -1 And Not MenuBar(2).IsSubMenu(MenuNumber) Then
        'HoverItem = MenuNumber
        If Index = 2 Then
            ShowStatus MenuBar(Index).ItemTooltipText(MenuNumber) & ". " & GetString(ResRightClickForHelp)
        Else
            ShowStatus MenuBar(Index).ItemTooltipText(MenuNumber) & ". "
        End If
        
        ShowToolTip MenuBar(Index).ItemTooltipText(MenuNumber)
    Else
        If ToolTipShown Then Unload frmToolTip
    End If

Case Else
    If MenuNumber > -1 Then
        ShowStatus MenuBar(Index).ItemTooltipText(MenuNumber) & ". "
    End If
    If ToolTipShown Then Unload frmToolTip
End Select
End Sub

Private Sub MenuBar_DragBar(Index As Integer, ByVal EventType As DragBarConstants, ByVal AuxInfo As Integer)
If ToolTipShown Then Unload frmToolTip
If EventType = AboutToPlace Then LockWindowUpdate FormMain.hWnd
If Visible And EventType = EndDrag Then
    Form_Resize
    LockWindowUpdate 0
    SaveSetting AppName, "General", "MenuAlign" & Index, Format(MenuBar(Index).Align)
End If
End Sub

Private Sub MenuBar_GotFocus(Index As Integer)
SetFocus
End Sub

Private Sub MenuBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
HideCooWhenMouseMoving
If MenuBar(Index).IsAboveDragArea(X, Y) And MenuBar(Index).DragAreaVisible Then
    ShowStatus GetString(ResClickToDragThisBar)
    MenuBar(Index).ToolTipText = ""
    'If ToolTipShown Then Unload frmToolTip
    Exit Sub
End If
If Button = 0 And MenuBar(Index).GetItemFromCursor(1, X, Y) = 0 Then
    ShowStatus
    MenuBar(Index).ToolTipText = ""
    'If ToolTipShown Then Unload frmToolTip
End If
End Sub

Private Sub MenuBar_MustHideTooltip(Index As Integer)
If ToolTipShown Then Unload frmToolTip
End Sub

Private Sub MenuBar_Resize(Index As Integer)
If Not FormResizeEngaged And Not MenubarResizeEngaged Then Form_Resize
End Sub

Private Sub MenuBar_WhatsThisRequest(Index As Integer, ByVal ItemIndex As Integer)
If ToolTipShown Then Unload frmToolTip
If Index = 1 Then
    MenuBar(Index).WhatsThisHelpID = ResHlp_PopBase + ItemIndex
End If
If Index = 2 Then
    Select Case MenuBar(Index).ToolbarMode
    Case mbsToolBar
        MenuBar(Index).WhatsThisHelpID = ResHlp_PopToolBase + Transpose(ItemIndex, MenuTransposition)
    Case mbsSelectObjectsFinish
        MenuBar(Index).WhatsThisHelpID = 9 + ItemIndex '?????
    End Select
End If
MenuBar(Index).ShowWhatsThis
End Sub

'===================================================
' Menu clicks processing
'===================================================

Private Sub mnuAbout_Click()
MenuCommand ResAbout
End Sub

Private Sub mnuActiveAxes_Click()
MenuCommand ResActiveAxes
End Sub

Private Sub mnuAnCircle_Click()
MenuCommand ResMnuAnCircle
End Sub

Private Sub mnuAnLine_Click()
MenuCommand ResMnuAnLine
End Sub

Private Sub mnuAnPoint_Click()
MenuCommand ResMnuAnPoint
End Sub

Private Sub mnuBezier_Click()
MenuCommand ResStaticObjectBase + sgBezier * 2
End Sub

Private Sub mnuBitmap_Click()
MenuCommand ResBMP
End Sub

Private Sub mnuButtonDelete_Click()
MenuCommand ResDeleteButton
End Sub

Private Sub mnuButtonMovable_Click()
MenuCommand ResFix + ResButton + 1
End Sub

Private Sub mnuButtonProperties_Click()
MenuCommand ResButtonProperties
End Sub

Private Sub mnuClearAll_Click()
MenuCommand ResClearAll
End Sub

Private Sub mnuClearLocus_Click()
MenuCommand ResDeleteLocus
End Sub

Private Sub mnuCreateLocus_Click()
MenuCommand ResCreateLocus
End Sub

Private Sub mnuDeleteFigure_Click()
MenuCommand ResMnuDeleteFigure
End Sub

Private Sub mnuDeleteLabel_Click()
MenuCommand ResMnuDeleteLabel
End Sub

Private Sub mnuDeletePoint_Click()
MenuCommand ResMnuDeletePoint
End Sub

Private Sub mnuDemo_Click()
MenuCommand ResDemo
End Sub

Private Sub mnuDemoOptions_Click()
MenuCommand ResDemoOptions
End Sub

Private Sub mnuDynamicLocus_Click()
MenuCommand ResFigureBase + 2 * dsDynamicLocus
End Sub

Private Sub mnuEdit_Click()
If Me.Fullscreen Then Exit Sub
mnuInsertButton.Enabled = GetDragsState = dscNormalState
End Sub

Private Sub mnuEnhancedMetafile_Click()
MenuCommand ResEMF
End Sub

Private Sub mnuExit_Click()
MenuCommand ResExit
End Sub

'Private Sub mnuExport_Click()
'MenuCommand mnuExport.Caption
'End Sub

Private Sub mnuFigureChoice_Click(Index As Integer)
ActiveFigure = mnuFigureChoice(Index).Tag
MenuCommand ResMnuChooseFigure
End Sub

Private Sub mnuFigureList_Click()
MenuCommand ResFigureList
End Sub

Private Sub mnuFigureProperties_Click()
MenuCommand ResMnuFigureProperties
End Sub

Private Sub mnuFileProps_Click()
MenuCommand ResFileProps
End Sub

Private Sub mnuFixLabel_Click()
MenuCommand ResFix + ResLabel + 1
End Sub

Private Sub mnuFullscreen_Click()
Me.Fullscreen = Not Me.Fullscreen
End Sub

Private Sub mnuHelpContents_Click()
MenuCommand ResHelpContents
End Sub

Private Sub mnuHideFigure_Click()
MenuCommand ResHide + ResFigure + 1
End Sub

Private Sub mnuHidePoint_Click()
MenuCommand ResHide + ResPoint + 1
End Sub

Private Sub mnuInsertButton_Click()
MenuCommand ResInsertButton
End Sub

Private Sub mnuInsertLabel_Click()
MenuCommand ResInsertLabel
End Sub

Private Sub mnuJSPHTML_Click()
MenuCommand ResJSPHTML
End Sub

Private Sub mnuLabelProperties_Click()
MenuCommand ResMnuLabelProperties
End Sub

Private Sub mnuLangEnglish_Click()
ChangeInterfaceLanguage langEnglish
End Sub

Private Sub mnuLangGerman_Click()
ChangeInterfaceLanguage langGerman
End Sub

Private Sub mnuLangRussian_Click()
ChangeInterfaceLanguage langRussian
End Sub

Private Sub mnuLangUkrainian_Click()
ChangeInterfaceLanguage langUkrainian
End Sub

Private Sub mnuMacroCreate_Click()
MenuCommand ResMnuMacroCreate
End Sub

Private Sub mnuMacroLoad_Click()
MenuCommand ResMnuMacroLoad
End Sub

Private Sub mnuMacroObject_Click(Index As Integer)
MacroObjectSelected mnuMacroObject(Index).Tag
End Sub

Private Sub mnuMacroResults_Click()
MenuCommand ResMnuMacroSelectResults
End Sub

Private Sub mnuMacroRun_Click(Index As Integer)
MacroRunMenu Index + 1
End Sub

Private Sub mnuMacroSave_Click()
MenuCommand ResMnuMacroSave
End Sub

Private Sub mnuMetafile_Click()
MenuCommand ResWMF
End Sub

'===================================================
'===================================================
'===================================================

Private Sub mnuMRUFile_Click(Index As Integer)
On Local Error GoTo EH
Dim S As String

S = MRUList(MRUList.Count + 1 - Index)
If Dir(S) = "" Then
    MsgBox GetString(ResMsgCannotOpenFile), vbCritical
    Exit Sub
End If

If IsDirty Then
    Select Case MsgBox(GetString(ResSave) & " " & DrawingName & "?", vbYesNoCancel + vbQuestion)
        Case vbYes
            MenuCommand ResSave
        Case vbNo
            'do nothing
        Case vbCancel
            Exit Sub
    End Select
End If
DrawingName = MRUList(MRUList.Count + 1 - Index)
AddMRUItem DrawingName
FormMain.ShowStatus GetString(ResWorkingPleaseWait)
FormMain.Enabled = False
Screen.MousePointer = vbHourglass
FormMain.Caption = GetString(ResCaption) + " - " + RetrieveName(DrawingName)
DoEvents
OpenFile DrawingName
FormMain.ShowStatus
Screen.MousePointer = vbDefault
FormMain.Enabled = True
If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
EH:
End Sub

'===================================================

Private Sub mnuNew_Click()
MenuCommand ResNew
End Sub

Private Sub mnuOpen_Click()
MenuCommand ResOpen
End Sub

Private Sub mnuPointChoice_Click(Index As Integer)
ActivePoint = mnuPointChoice(Index).Tag
MenuCommand ResMnuChoosePoint
End Sub

Private Sub mnuPointList_Click()
MenuCommand ResPointList
End Sub

Private Sub mnuPointPopup_Click()
mnuClearLocus.Enabled = BasePoint(ActivePoint).Locus <> 0
End Sub

Private Sub mnuPointProperties_Click()
MenuCommand ResMnuPointProperties
End Sub

Private Sub mnuPolygon_Click()
'MenuCommand ResStaticObjectBase + 2 * sgPolygon
SelectTool dsPolygon
End Sub

Private Sub mnuPrint_Click()
MenuCommand ResPrint
End Sub

Private Sub mnuRecalcLabel_Click()
MenuCommand ResMnuRecalcLabel
End Sub

Private Sub mnuRedo_Click()
MenuCommand ResRedo
End Sub

Private Sub mnuReleasePoint_Click()
MenuCommand ResMnuReleasePoint
End Sub

Private Sub mnuSave_Click()
MenuCommand ResSave
End Sub

Private Sub mnuSaveAs_Click()
MenuCommand ResSaveAs
End Sub

Private Sub mnuSettings_Click()
MenuCommand ResOptions
End Sub

Private Sub mnuSGDelete_Click()
MenuCommand ResStaticObjectBase + 1 'static delete sign for MenuCommand
End Sub

Private Sub mnuSGProperties_Click()
MenuCommand ResStaticObjectBase + 3 'properties sign for MenuCommand
End Sub

Public Sub mnuShowAxes_Click()
MenuCommand ResShowAxes
End Sub

Public Sub mnuShowGrid_Click()
MenuCommand ResShowGrid
End Sub

Private Sub mnuShowPointName_Click()
MenuCommand ResShowName
End Sub

Private Sub mnuShowRulers_Click()
ToggleRulers
End Sub

Private Sub mnuShowStatusbar_Click()
ToggleStatusbar
End Sub

Private Sub mnuShowToolbar_Click()
ToggleToolbar
End Sub

Private Sub mnuShowMainbar_Click()
ToggleMainbar
End Sub

Private Sub mnuSnapTo_Click(Index As Integer)
ActiveFigure = mnuSnapTo(Index).Tag
MenuCommand ResMnuSnapToFigure
End Sub

Private Sub mnuTip_Click()
MenuCommand ResTipOfTheDay
End Sub

'===================================================
' Now tool menus
'===================================================

Private Sub mnuToolArc_Click()
SelectTool dsCircle_ArcCenterAndRadiusAndTwoPoints
End Sub

Private Sub mnuToolBisector_Click()
SelectTool dsBisector
End Sub

Private Sub mnuToolCircle_Click()
SelectTool dsCircle_CenterAndCircumPoint
End Sub

Private Sub mnuToolCircleByRadius_Click()
SelectTool dsCircle_CenterAndTwoPoints
End Sub

Private Sub mnuToolIntersect_Click()
SelectTool dsIntersect
End Sub

Private Sub mnuToolInvert_Click()
SelectTool dsInvert
End Sub

Private Sub mnuToolLine_Click()
SelectTool dsLine_2Points
End Sub

Private Sub mnuToolMeasureAngle_Click()
SelectTool dsMeasureAngle
End Sub

Private Sub mnuToolMeasureArea_Click()
SelectTool dsMeasureArea
End Sub

Private Sub mnuToolMeasureDistance_Click()
SelectTool dsMeasureDistance
End Sub

Private Sub mnuToolMiddlePoint_Click()
SelectTool dsMiddlePoint
End Sub

Private Sub mnuToolParallelLine_Click()
SelectTool dsLine_PointAndParallelLine
End Sub

Private Sub mnuToolPerpendicularLine_Click()
SelectTool dsLine_PointAndPerpendicularLine
End Sub

Private Sub mnuToolPoint_Click()
SelectTool dsPoint
End Sub

Private Sub mnuToolPointOnFigure_Click()
SelectTool dsPointOnFigure
End Sub

Private Sub mnuToolRay_Click()
SelectTool dsRay
End Sub

Private Sub mnuToolReflectedPoint_Click()
SelectTool dsSimmPointByLine
End Sub

Private Sub mnuToolSegment_Click()
SelectTool dsSegment
End Sub

Private Sub mnuToolSymmPoint_Click()
SelectTool dsSimmPoint
End Sub

Private Sub mnuUndo_Click()
MenuCommand ResUndo
End Sub

Private Sub mnuUndoBufferSize_Click()
MsgBox "Undo memory size is " & UndoMemoryConsumption & " bytes", vbInformation
End Sub

'===================================================
'===================================================

Private Sub mnuVector_Click()
MenuCommand ResStaticObjectBase + 2 * sgVector
End Sub

Private Sub mnuVectorDel_Click(Index As Integer)
ActiveStatic = mnuVectorDel(Index).Tag
MenuCommand ResStaticObjectBase + 1 'deletion of static objects sign for MenuCommand
End Sub

Private Sub mnuVectorProp_Click(Index As Integer)
ActiveStatic = mnuVectorProp(Index).Tag
frmStaticProps.Show
End Sub

Private Sub mnuView_Click()
mnuShowAxes.Checked = nShowAxes
mnuShowGrid.Checked = nShowGrid
mnuShowRulers.Checked = setShowRulers
End Sub

Private Sub mnuWE_Click()
MenuCommand mnuWE.Caption
End Sub

'=========================================================
'End of menu message map
'=========================================================


Private Sub Status_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 2 Then
    ShowStatus GetString(ResCursorCoordinates) & "."
Else
    If Button = 0 Then ShowStatus ""
End If
End Sub

Private Sub StatusBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then ShowStatus ""
End Sub

Public Sub StatusBar_Resize()
On Local Error Resume Next
Dim CoordWidth As String, LeftEdge As Long, RightEdge As Long, StatusHeight As Long

If StatusBarResizeEngaged Then Exit Sub
StatusBarResizeEngaged = True

If StatusBarAutoRedraw Then If setGradientFill Then Gradient StatusBar.hDC, StatusBar.BackColor, setcolToolbar, 0, 0, StatusBar.ScaleWidth, StatusBar.ScaleHeight Else StatusBar.Cls

StatusHeight = StatusBar.TextHeight("W")
Status(1).Move Margin, (StatusBar.ScaleHeight - StatusHeight) \ 2
If setShowClock Then Clock.Move StatusBar.ScaleWidth - Clock.Width - Margin, (StatusBar.ScaleHeight - StatusHeight) \ 2
If setShowCoord Then
    CoordWidth = "X =-" & setFormatDistance & ";" & " Y =-" & setFormatDistance
    If Status(2).Caption = "" Then Status(2).Caption = CoordWidth
    
    LeftEdge = Margin
    RightEdge = StatusBar.ScaleWidth - Margin
    If setShowClock Then RightEdge = RightEdge - Clock.Width - Margin
    
    Status(2).Move RightEdge - 2 * Margin - StatusBar.TextWidth(CoordWidth), (StatusBar.ScaleHeight - StatusHeight) \ 2, StatusBar.TextWidth(CoordWidth) + 4
    If Status(2).Left - 2 * Margin > 0 Then Status(1).Move Margin, Status(2).Top, Status(2).Left - 2 * Margin, StatusHeight
End If

StatusBarResizeEngaged = False
End Sub

Public Sub UpdateCoords(ByVal X As Double, ByVal Y As Double)
If setShowCoord And Not StatusBarSpecialMode Then
    Status(2).Caption = "X =" & Format(X, IIf(X >= 0, " ", "") & setFormatDistance) & ";" & " Y =" & Format(Y, IIf(Y >= 0, " ", "") & setFormatDistance)
    If StatusBarAutoRedraw Then Status(2).Refresh
    If Not Status(2).Visible Then Status(2).Visible = True
End If
End Sub

Public Sub SetMousePointer(ByVal MousePointer As CursorState)
If Not setLoadCursors Then
    If Paper.MousePointer <> vbDefault Then Paper.MousePointer = vbDefault
    Exit Sub
Else
    If Paper.MousePointer = vbDefault Then Paper.MousePointer = setPaperCursorArrow
End If
If PaperCursor = MousePointer Then Exit Sub
PaperCursor = MousePointer
Select Case MousePointer
    Case curStateArrow
        Set Paper.MouseIcon = curArrow
        If Paper.MousePointer <> setPaperCursorArrow Then Paper.MousePointer = setPaperCursorArrow
    Case curStateCross
        Set Paper.MouseIcon = curArrowCross
        If Paper.MousePointer <> setPaperCursorCross Then Paper.MousePointer = setPaperCursorCross
    Case curStateDrag
        Set Paper.MouseIcon = curArrowDrag
        If Paper.MousePointer <> setPaperCursorDrag Then Paper.MousePointer = setPaperCursorDrag
    Case curStateNo
        If Paper.MousePointer <> vbNoDrop Then Paper.MousePointer = vbNoDrop
    Case curStateQuestion
        If Paper.MousePointer <> vbArrowQuestion Then Paper.MousePointer = vbArrowQuestion
    Case curStateHourglass
        If Paper.MousePointer <> vbHourglass Then Paper.MousePointer = vbHourglass
    Case curStateAdd
        Set Paper.MouseIcon = curArrowPlus
        If Paper.MousePointer <> setPaperCursorArrow Then Paper.MousePointer = setPaperCursorArrow
    Case curStateRemove
        Set Paper.MouseIcon = curArrowMinus
        If Paper.MousePointer <> setPaperCursorArrow Then Paper.MousePointer = setPaperCursorArrow
    Case curStateSelect
        Set Paper.MouseIcon = curArrowSelect
        If Paper.MousePointer <> setPaperCursorArrow Then Paper.MousePointer = setPaperCursorArrow
End Select
'If Paper.MousePointer <> MousePointer Then Paper.MousePointer = MousePointer
End Sub

Public Sub CancelMacroRun()
i_ExitMacroRunMode
End Sub

Public Sub ShowStatus(Optional ByVal SString As String, Optional ByVal Num As Integer = 1)
Dim RightBorder As Long, sStr As String
If Status(Num).Caption = SString Or StatusBarSpecialMode Then Exit Sub
On Local Error GoTo EH

If Num = 1 Then
    If Status(2).Visible Then
        RightBorder = Status(2).Left - 2 * Margin
    Else
        If setShowClock Then RightBorder = Clock.Left - 2 * Margin Else RightBorder = StatusBar.ScaleWidth - 2 * Margin
    End If
    Do While StatusBar.TextWidth(SString) >= RightBorder And SString <> "..."
        If Len(SString) < 4 Then Exit Do
        SString = Left(SString, Len(SString) - 4) + "..."
    Loop
End If
If Num = 2 Then
    RightBorder = Status(2).Left - 2 * Margin
    sStr = Status(1).Caption
    Do While StatusBar.TextWidth(sStr) >= RightBorder And sStr <> "..."
        If Len(sStr) < 4 Then Exit Do
        sStr = Left(sStr, Len(sStr) - 4) + "..."
    Loop
    If sStr <> Status(1).Caption Then Status(1).Caption = sStr: Status(1).Refresh
End If
Status(Num).Caption = SString
If StatusBarAutoRedraw Then Status(Num).Refresh

EH:
End Sub

Public Sub ShowStatusSpecial(Optional ByVal SString As String)
Dim RightBorder As Long, sStr As String
If Status(1).Caption = SString Then Exit Sub
On Local Error GoTo EH

If Status(2).Visible Then
    RightBorder = Status(2).Left - 2 * Margin
Else
    If setShowClock Then RightBorder = Clock.Left - 2 * Margin Else RightBorder = StatusBar.ScaleWidth - 2 * Margin
End If
Do While StatusBar.TextWidth(SString) >= RightBorder And SString <> "..."
    If Len(SString) < 4 Then Exit Do
    SString = Left(SString, Len(SString) - 4) + "..."
Loop

Status(1).Caption = SString
If StatusBarAutoRedraw Then Status(1).Refresh

'Const AnimSteps = 1000
'Dim Z As Long, R As Long, G As Long, B As Long
'Dim DR As Long, DG As Long, DB As Long
'Dim SC As Long, DC As Long
'SC = vbWhite
'DC = FormMain.Status(1).ForeColor
'If DC < 0 Then DC = GetSysColor(DC + SysColorTranslationBase)
'R = Red(SC)
'G = Green(SC)
'B = Blue(SC)
'DR = Red(DC)
'DG = Green(DC)
'DB = Blue(DC)
'For Z = 1 To AnimSteps
'    FormMain.Status(1).ForeColor = RGB(R + (DR - R) * Z / AnimSteps, G + (DG - G) * Z / AnimSteps, B + (DB - B) * Z / AnimSteps)
'    Status(1).Refresh
'Next

EH:
End Sub

'###############################################################
'Ruler-related stuff
'###############################################################

Public Sub MoveIndicX(ByVal X As Double)
If Not setShowRulers Then Exit Sub
If Not setShowCoord Then
    If IndicX.Visible Then IndicX.Visible = False
    Exit Sub
End If
If Not IndicX.Visible Then IndicX.Visible = True
IndicX.X1 = X
IndicX.X2 = X
IndicX.Y1 = 0
IndicX.Y2 = XRuler.ScaleHeight
End Sub

Public Sub MoveIndicY(ByVal Y As Double)
If Not setShowRulers Then Exit Sub
If Not setShowCoord Then
    If IndicY.Visible Then IndicY.Visible = False
    Exit Sub
End If
If Not IndicY.Visible Then IndicY.Visible = True
IndicY.X1 = 0
IndicY.X2 = YRuler.ScaleWidth
IndicY.Y1 = Y
IndicY.Y2 = Y
End Sub

Public Sub HideIndicX()
IndicX.Visible = False
End Sub

Public Sub HideIndicY()
IndicY.Visible = False
End Sub

Private Sub XRuler_MouseMove(Button As Integer, Shift As Integer, OX As Single, OY As Single)
HideIndicX
HideIndicY
Status(2).Visible = False
End Sub

Public Sub XRuler_Resize()
Dim X As Long, Y As Long, tX As Double, tY As Double, Q As Double, SH As Long, DC As Long
Dim LF As LOGFONT, i As Long, NF As Long, tS As String, TI As Long, lpSize As Size
If CanvasBorders.P1.X >= CanvasBorders.P2.X Or Not setShowRulers Or Not XRuler.Visible Then Exit Sub

With XRuler
    'If XRuler.BackColor <> setcolRuler Then XRuler.BackColor = setcolRuler
    If setGradientFill Then Gradient .hDC, colRulerGradient, .BackColor, 0, 0, .ScaleWidth, .ScaleHeight Else .Cls
    SH = XRuler.ScaleHeight - 1
    DC = XRuler.hDC
    
    For Q = Round(CanvasBorders.P1.X, 1) To Round(CanvasBorders.P2.X, 1) Step 0.1
        tX = Q
        tY = 0
        ToPhysical tX, tY
        SetPixelV DC, tX, SH, 0
        SetPixelV DC, tX, SH - 1, 0
    Next
    For X = CanvasBorders.P1.X To CanvasBorders.P2.X
        tX = X
        tY = 0
        ToPhysical tX, tY
        SetPixelV DC, tX, SH - 2, 0
    Next
    
    LF.lfWidth = 0
    LF.lfEscapement = 0
    LF.lfOrientation = 0
    LF.lfWeight = 400
    LF.lfItalic = 0
    LF.lfUnderline = 0
    LF.lfStrikeOut = 0
    LF.lfCharSet = 0 'Paper.Font.Charset
    LF.lfOutPrecision = 0
    LF.lfClipPrecision = 0
    LF.lfQuality = 2
    LF.lfPitchAndFamily = 0
    tS = RulerFontName
    'LF.lfFaceName = tS & vbNullChar '?????
    For Q = 1 To Len(tS)
        LF.lfFaceName(Q - 1) = Asc(Mid$(tS, Q, 1))
    Next
    LF.lfFaceName(Len(tS)) = 0
    LF.lfHeight = RulerFontSize * -20 / Screen.TwipsPerPixelY
    
    i = CreateFontIndirect(LF)
    NF = SelectObject(DC, i)
    SetTextAlign DC, TA_BOTTOM Or TA_CENTER
    
    GetTextExtentPoint32 DC, "0" & vbNullChar, 1, lpSize
    Y = lpSize.cy
    
    TI = -21
    For X = Round(CanvasBorders.P1.X) To Round(CanvasBorders.P2.X)
        tX = X
        tY = 0
        ToPhysical tX, tY
        If tX > TI + 20 Then
            TextOut DC, tX, Y, CStr(X), Len(CStr(X))
            TI = tX
        End If
    Next
    
    SelectObject DC, NF
    DeleteObject i
    .Refresh
End With
'    .ForeColor = colRulerForeColor
'    .FontSize = RulerFontSize
'    Y = .TextHeight("0")
'    For X = Round(CanvasBorders.P1.X) To Round(CanvasBorders.P2.X)
'        TX = X
'        TY = 0
'        ToPhysical TX, TY
'        TextOut DC, TX, Y, CStr(X), Len(CStr(X)) 'XRuler, X, TX, Y
'    Next
   
End Sub

Private Sub YRuler_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HideIndicX
HideIndicY
Status(2).Visible = False
End Sub

Public Sub YRuler_Resize()
Dim X As Long, Y As Long, tX As Double, tY As Double, SW As Long, SH As Long, tTH As Long, Q As Double
Dim LF As LOGFONT, i As Long, NF As Long, tS As String, DC As Long, TI As Long, lpSize As Size

If CanvasBorders.P1.Y >= CanvasBorders.P2.Y Or Not setShowRulers Or Not YRuler.Visible Then Exit Sub

With YRuler
    'If YRuler.BackColor <> setcolRuler Then YRuler.BackColor = setcolRuler
    If setGradientFill Then Gradient .hDC, colRulerGradient, .BackColor, 0, 0, .ScaleWidth, .ScaleHeight, False Else .Cls
    SW = YRuler.ScaleWidth - 1
    DC = .hDC
    For Q = Round(CanvasBorders.P1.Y, 1) To Round(CanvasBorders.P2.Y, 1) Step 0.1
        tX = 0
        tY = Q
        ToPhysical tX, tY
        SetPixelV DC, SW, tY, 0
        SetPixelV DC, SW - 1, tY, 0
    Next
    For Y = CanvasBorders.P1.Y To CanvasBorders.P2.Y
        tX = 0
        tY = Y
        ToPhysical tX, tY
        SetPixelV DC, SW - 2, tY, 0
    Next

    LF.lfWidth = 0
    LF.lfEscapement = 900
    LF.lfOrientation = 900
    LF.lfWeight = 400
    LF.lfItalic = 0
    LF.lfUnderline = 0
    LF.lfStrikeOut = 0
    LF.lfCharSet = 0 'Paper.Font.Charset
    LF.lfOutPrecision = 0
    LF.lfClipPrecision = 0
    LF.lfQuality = 2
    LF.lfPitchAndFamily = 0
    tS = RulerFontName
    'LF.lfFaceName = tS & vbNullChar
    For Q = 1 To Len(tS)
        LF.lfFaceName(Q - 1) = Asc(Mid$(tS, Q, 1))
    Next
    LF.lfFaceName(Len(tS)) = 0
    LF.lfHeight = RulerFontSize * -20 / Screen.TwipsPerPixelY
    
    i = CreateFontIndirect(LF)
    NF = SelectObject(DC, i)
    SetTextAlign DC, TA_BOTTOM Or TA_CENTER
    
    SW = .ScaleWidth
    SH = .ScaleHeight
    'tTH = .TextHeight("0") / 2
    GetTextExtentPoint32 DC, "0" & vbNullChar, 1, lpSize
    tTH = lpSize.cy / 2
    X = SW - 2
    TI = 100000
    For Y = Round(CanvasBorders.P1.Y) To Round(CanvasBorders.P2.Y)
        tX = 0
        tY = Y
        ToPhysical tX, tY
        If tY < TI - 20 Then
            TextOut DC, X, tY, Y, Len(CStr(Y))
            TI = tY
        End If
    Next
    
    SelectObject DC, NF
    DeleteObject i

    .Refresh
End With
End Sub

'##############################################################
'End of ruler-related stuff
'##############################################################

Private Sub tmrDemo_Timer()
If DragS.State = dscDemo Then
    If DemoSequence.Step = DemoSequence.StepCount Then
        DemoFirstStep
    Else
        DemoNextStep
    End If
End If
End Sub

Private Sub ValueTable1_NeedToClose()
FormMain.Docked.Visible = False
FormMain.Form_Resize
FormMain.mnuWE.Checked = False
End Sub

Private Sub ValueTable1_NeedToHide()
OldDockedWidth = Docked.Width
Docked.Width = 4
Docked.MousePointer = vbArrow
Form_Resize
End Sub

Private Sub ValueTable1_NeedToShow()
Docked.Width = OldDockedWidth
OldDockedWidth = 0
Docked.MousePointer = vbSizeWE
Form_Resize
End Sub

Private Sub DrawStripe(ByVal X As Long)
Dim lpRect As RECT, DC As Long, Z As Long
Const StripeWidth = 2
GetWindowRect Docked.hWnd, lpRect
lpRect.Left = lpRect.Left + X - StripeWidth
lpRect.Right = lpRect.Left + 2 * StripeWidth
DC = GetWindowDC(GetDesktopWindow)
For Z = 1 To StripeWidth
    DrawFocusRect DC, lpRect
    lpRect.Left = lpRect.Left + 1
    lpRect.Top = lpRect.Top + 1
    lpRect.Right = lpRect.Right - 1
    lpRect.Bottom = lpRect.Bottom - 1
Next Z
End Sub

Public Property Get Fullscreen() As Boolean
Fullscreen = IsFullScreen
End Property

Public Property Let Fullscreen(ByVal vNewValue As Boolean)
Static WasToolbarShown As Boolean
Static WasWEShown As Boolean
Static WasStatusbarShown As Boolean
Static WasRulerShown As Boolean
Static WasWindowMaximized As Boolean
Static WasMainbarShown As Boolean
Static OldWindowState As Integer

If vNewValue = IsFullScreen Then Exit Property

LockWindowUpdate Me.hWnd
FormResizeEngaged = True

IsFullScreen = vNewValue
If IsFullScreen Then
    WasToolbarShown = mnuShowToolbar.Checked
    WasStatusbarShown = mnuShowStatusbar.Checked
    WasRulerShown = mnuShowRulers.Checked
    WasWEShown = mnuWE.Checked
    WasWindowMaximized = Me.WindowState = vbMaximized
    WasMainbarShown = mnuShowMainbar.Checked
    
    If WasWEShown Then mnuWE_Click
    If WasToolbarShown Then ToggleToolbar False
    If WasMainbarShown Then ToggleMainbar False
    If WasStatusbarShown Then ToggleStatusbar False
    If WasRulerShown Then ToggleRulers False
    OldWindowState = Me.WindowState
    If Not WasWindowMaximized Then Me.WindowState = vbMaximized
    'EnableMenu False, , False, True
    EnableMenus mnsFullScreen, False
    MakeFullScreen hWnd, vNewValue
Else
    MakeFullScreen hWnd, vNewValue
    If WasWEShown Then mnuWE_Click
    If WasToolbarShown Then ToggleToolbar False
    If WasMainbarShown Then ToggleMainbar False
    If WasStatusbarShown Then ToggleStatusbar False
    If WasRulerShown Then ToggleRulers False
    If Not WasWindowMaximized Then Me.WindowState = OldWindowState
    If Me.Caption <> GetString(ResCaption) + " - " + RetrieveName(DrawingName) Then FormMain.Caption = GetString(ResCaption) + " - " + RetrieveName(DrawingName)
    'EnableMenu , , False, True
    EnableMenus mnsStandard
End If
mnuFullscreen.Checked = vNewValue

FormResizeEngaged = False
Form_Resize

LockWindowUpdate 0
End Property

Public Sub PrepareFormControls(Optional ByVal ShouldSendResize As Boolean = False)
PrepareToolbar
PrepareRulers ShouldSendResize
PrepareStatusBar ShouldSendResize
End Sub

Public Sub PrepareStatusBar(Optional ByVal ShouldSendResize As Boolean = True)
Dim StatusHeight As Long
Status(1).Font.Name = StatusBar.Font.Name
Status(2).Font.Name = StatusBar.Font.Name
Status(1).Font.Size = StatusBar.Font.Size
Status(2).Font.Size = StatusBar.Font.Size
If StatusBar.Visible <> setShowStatusbar Then StatusBar.Visible = setShowStatusbar
If Status(2).Visible <> setShowCoord Then Status(2).Visible = setShowCoord
If Clock.Visible <> setShowClock Then Clock.Visible = setShowClock
StatusHeight = StatusBar.TextHeight("W")
If StatusBar.Height <> StatusHeight + 6 Then StatusBar.Height = StatusHeight + 6
mnuShowStatusbar.Checked = setShowStatusbar
If ShouldSendResize Then StatusBar_Resize
End Sub

Public Sub PrepareToolbar()
Dim Z As Long, Q As Long

If MenuBar(1).Visible <> setShowMainbar Then MenuBar(1).Visible = setShowMainbar
If MenuBar(2).Visible <> setShowToolbar Then MenuBar(2).Visible = setShowToolbar
For Z = MenuBar.LBound To MenuBar.UBound
    If MenuBar(Z).GradientColor <> setcolToolbar Then MenuBar(Z).GradientColor = setcolToolbar
    If MenuBar(Z).Gradiented <> setGradientFill Then MenuBar(Z).Gradiented = setGradientFill
    Q = Val(GetSetting(AppName, "General", "MenuAlign" & Z, Format(Choose(Z, Format(DefaultAlign), Format(DefaultAlign)))))
    If MenuBar(Z).Align <> Q Then MenuBar(Z).Align = Q
Next
mnuShowToolbar.Checked = setShowToolbar
mnuShowMainbar.Checked = setShowMainbar
End Sub

Public Sub PrepareRulers(Optional ByVal ShouldSendResize As Boolean = False)
If XRuler.BackColor <> setcolRuler Then XRuler.BackColor = setcolRuler
If YRuler.BackColor <> setcolRuler Then YRuler.BackColor = setcolRuler
If setShowRulers Then
    If YRuler.FontSize <> RulerFontSize Then YRuler.FontSize = RulerFontSize
    If YRuler.Width <> YRuler.TextHeight("0") + 4 Then YRuler.Width = YRuler.TextHeight("0") + 4
    If XRuler.FontSize <> RulerFontSize Then XRuler.FontSize = RulerFontSize
    If XRuler.Height <> XRuler.TextHeight("0") + 4 Then XRuler.Height = XRuler.TextHeight("0") + 4
    If Not XRuler.Visible Then XRuler.Visible = True
    If Not YRuler.Visible Then YRuler.Visible = True
    If Not RulerButton.Visible Then RulerButton.Visible = True
Else
    If XRuler.Visible Then XRuler.Visible = False
    If YRuler.Visible Then YRuler.Visible = False
    If RulerButton.Visible Then RulerButton.Visible = False
End If
If ShouldSendResize Then
    XRuler_Resize
    YRuler_Resize
End If
mnuShowRulers.Checked = setShowRulers
End Sub

Public Sub EnableUndoMenu()
mnuUndo.Enabled = DragS.State = dscNormalState
End Sub

'=========================================================
'                                           DEMO MODE
'=========================================================

Public Sub EnterObjectSelectionMode(ByVal oType As ObjectSelectionType, Optional ByVal cType As ObjectSelectionCaller = oscButton)
Dim TempMBS As MenuBarState
TempMBS = mbsSelectObjectsFinish
Select Case oType
Case ostShowHideObjects
    DragS.State = dscSelectObjects
Case ostCalcPoints
    DragS.State = dscSelectObjects
    If TempObjectSelection.PointCountMax > 0 Then TempMBS = mbsCancel
End Select

FillMenuBar TempMBS
MenuBar(1).Enabled = False

DrawingState = dsSelect
'EnableMenu False, , False
EnableMenus mnsObjectSelection
ShowStatus GetString(ResSelectObjects)
PaperCls
ShowSelectedAll TempObjectSelection
FormMain.Enabled = True
FormMain.SetFocus
End Sub

Public Sub ExitObjectSelectionMode()
DragS.State = dscNormalState
ShowStatus
FillMenuBar
MenuBar(1).Enabled = True
'EnableMenu True, , False
EnableMenus mnsStandard
PaperCls
ShowAll
End Sub

'=========================================================
'                                           DEMO MODE
'=========================================================

Public Sub EnterDemoMode()
DragS.State = dscDemo
FillMenuBar mbsDemo
MenuBar(1).Enabled = False
EnterStatusbarSpecialMode
FormMain.Caption = GetString(ResCaption) + " - " + RetrieveName(DrawingName) & " - " & GetString(ResDemo)
If nDemoInterval < 500 Or nDemoInterval > 15000 Then nDemoInterval = defDemoInterval
tmrDemo.Interval = nDemoInterval

DrawingState = dsSelect
'EnableMenu False, , False
EnableMenus mnsDemo
ShowStatusSpecial GetString(ResStepFirst) & "."
PaperCls
DemoFirstStep
End Sub

Public Sub ExitDemoMode()
tmrDemo.Enabled = False
DragS.State = dscNormalState
ShowStatus
FillMenuBar
MenuBar(1).Enabled = True
ExitStatusbarSpecialMode
'EnableMenu True, , False
EnableMenus mnsStandard
FormMain.Caption = GetString(ResCaption) + " - " + RetrieveName(DrawingName)
PaperCls
ShowAll
End Sub

'========================================================
'                               Macro Givens Select Mode
'========================================================

Public Sub EnterMacroGivenSelectMode()
DragS.State = dscMacroStateGivens
FillMenuBar mbsMacroGivens
EnableMenus mnsMacroGivens
End Sub

Public Sub ExitMacroCreateMode()
DragS.State = dscNormalState
FillMenuBar mbsToolBar
EnableMenus mnsStandard
mnuMacroSave.Visible = False
mnuMacroResults.Visible = False
End Sub

'========================================================
'                               Macro Results Select Mode
'========================================================

Public Sub EnterMacroResultSelectMode()
DragS.State = dscMacroStateResults
FillMenuBar mbsMacroResults
mnuMacroSave.Visible = True
mnuMacroResults.Visible = False
End Sub

'========================================================
'                               Macro Run Mode
'========================================================

Public Sub EnterMacroRunMode()
DragS.State = dscMacroStateRun
EnableMenus mnsMacroRun
FillMenuBar mbsMacroRun
End Sub

Public Sub ExitMacroRunMode()
Dim Z As Long

ReDim DragS.MacroObjectDescription(1 To 1)
ReDim DragS.MacroObjects(1 To 1)
ReDim DragS.MacroObjectType(1 To 1)
DragS.MacroObjectCount = 0
DragS.MacroCurrentObject = 1
For Z = 1 To MaxDragNumbers: DragS.Number(Z) = 0:  Next
DragS.State = dscNormalState

FillMenuBar mbsToolBar
EnableMenus mnsStandard

PaperCls
ShowAll
ImitateMouseMove
End Sub

'========================================================
'                               Statusbar special mode
'========================================================

Public Sub EnterStatusbarSpecialMode()
StatusBarSpecialMode = True
Status(2).Visible = False
End Sub

Public Sub ExitStatusbarSpecialMode()
StatusBarSpecialMode = False
PrepareStatusBar
ShowStatus
End Sub

'========================================================
'                           Auxiliary properties and methods
'========================================================

Public Function GetDragsState() As DragStateConstants
GetDragsState = DragS.State
End Function

Public Sub UpdateRulers(Optional ByVal BitField As Long = 3)
If XRuler.Visible Then If (BitField And 1) = 1 Then XRuler_Resize
If YRuler.Visible Then If (BitField And 2) = 2 Then YRuler_Resize
End Sub

Public Sub FillStrings()
DrawingName = GetString(ResUntitled)
FormMain.Caption = GetString(ResCaption) + " - " + DrawingName
FormMain.Status(2).ToolTipText = GetString(ResCurrentTool)
'FormMain.ValueTable1.Refresh
End Sub

Public Sub ChangeInterfaceLanguage(ByVal NewLang As Languages)
LockWindowUpdate FormMain.hWnd
setLanguage = NewLang
SaveSetting AppName, "General", "Language", Format(setLanguage)
FillStrings
CreateMenus
DoEvents
LockWindowUpdate 0
End Sub

'========================================================
'                          Toggle on/off some of the interface elements
'========================================================

Public Sub ToggleRulers(Optional ByVal ShouldLockWindow As Boolean = True, Optional ByVal ShouldToggle As Boolean = True, Optional ByVal ShouldSaveSetting As Boolean = True)
If ShouldToggle Then setShowRulers = Not setShowRulers
mnuShowRulers.Checked = setShowRulers
If ShouldSaveSetting Then SaveSetting AppName, "Interface", "ShowRulers", Format(-CInt(setShowRulers))
If ShouldLockWindow Then LockWindowUpdate Me.hWnd

If setShowRulers Then
    XRuler.Visible = True
    YRuler.Visible = True
    RulerButton.Visible = True
    If YRuler.FontSize <> RulerFontSize Then YRuler.FontSize = RulerFontSize
    If YRuler.Width <> YRuler.TextHeight("0") + 4 Then YRuler.Width = YRuler.TextHeight("0") + 4
    If XRuler.FontSize <> RulerFontSize Then XRuler.FontSize = RulerFontSize
    If XRuler.Height <> XRuler.TextHeight("0") + 4 Then XRuler.Height = XRuler.TextHeight("0") + 4
Else
    YRuler.Visible = False
    XRuler.Visible = False
    RulerButton.Visible = False
End If

Form_Resize
PaperCls
ShowAll
If ShouldLockWindow Then LockWindowUpdate 0
End Sub

Public Sub ToggleStatusbar(Optional ByVal ShouldLockWindow As Boolean = True)
setShowStatusbar = Not setShowStatusbar
SaveSetting AppName, "Interface", "ShowStatusbar", setShowStatusbar
mnuShowStatusbar.Checked = setShowStatusbar
If ShouldLockWindow Then LockWindowUpdate Me.hWnd
StatusBar.Visible = setShowStatusbar
Form_Resize
If ShouldLockWindow Then LockWindowUpdate 0
End Sub

Public Sub ToggleToolbar(Optional ByVal ShouldLockWindow As Boolean = True)
setShowToolbar = Not setShowToolbar
SaveSetting AppName, "Interface", "ShowToolbar", setShowToolbar
mnuShowToolbar.Checked = setShowToolbar
If ShouldLockWindow Then LockWindowUpdate Me.hWnd
MenuBar(2).Visible = setShowToolbar
Form_Resize
If ShouldLockWindow Then LockWindowUpdate 0
End Sub

Public Sub ToggleMainbar(Optional ByVal ShouldLockWindow As Boolean = True)
setShowMainbar = Not setShowMainbar
SaveSetting AppName, "Interface", "ShowMainbar", setShowMainbar
mnuShowMainbar.Checked = setShowMainbar
If ShouldLockWindow Then LockWindowUpdate Me.hWnd
MenuBar(1).Visible = setShowMainbar
Form_Resize
If ShouldLockWindow Then LockWindowUpdate 0
End Sub

Public Sub EnableMenu(Optional ByVal Enable As Boolean = True, Optional ByVal EnableMacros As Integer = 3, Optional ByVal ShouldTouchToolbar As Boolean = True, Optional ByVal FromFullscreen As Boolean = False)

'Dim Z As Long
'If ShouldTouchToolbar Then
'    For Z = MenuBar.LBound To MenuBar.UBound
'        MenuBar(Z).Enabled = Enable
'    Next
'End If
'
'Docked.Enabled = Enable
'
'EFile:
'mnuNew.Enabled = Enable
'mnuOpen.Enabled = Enable
'mnuSave.Enabled = Enable
'mnuSaveAs.Enabled = Enable
'mnuPrint.Enabled = Enable
'mnuExport.Enabled = Enable
'If Not FromFullscreen Then mnuExit.Enabled = Enable
'For Z = 1 To mnuMRUFile.UBound
'    mnuMRUFile(Z).Enabled = Enable
'Next
'
'EEdit:
'mnuUndo.Enabled = Enable
'mnuRedo.Enabled = Enable
'If Enable Then
'    mnuUndo.Enabled = ActivityCount > 0
'    mnuRedo.Enabled = UndoneActivityCount > 0
'End If
'mnuInsertLabel.Enabled = Enable
'mnuInsertButton.Enabled = Enable
'mnuClearAll.Enabled = Enable
'If Not FromFullscreen Then mnuFileProps.Enabled = Enable
'
'EView:
'mnuShowAxes.Enabled = Enable
'mnuShowGrid.Enabled = Enable
'mnuShowRulers.Enabled = Enable
'mnuShowStatusbar.Enabled = Enable
'mnuShowToolbar.Enabled = Enable
'mnuShowMainbar.Enabled = Enable
'mnuWE.Enabled = Enable
'mnuPointList.Enabled = Enable
'mnuFigureList.Enabled = Enable
'If Not FromFullscreen Then mnuFullscreen.Enabled = Enable
'mnuDemo.Enabled = Enable
'mnuDemoOptions.Enabled = Enable
'
'EFigures:
'mnuAnalytic.Enabled = Enable
'mnuFigCircles.Enabled = Enable
'mnuFigConstruction.Enabled = Enable
'mnuFigLines.Enabled = Enable
'mnuFigMeasure.Enabled = Enable
'mnuFigPoints.Enabled = Enable
'
'EOptions:
'mnuSettings.Enabled = Enable
'mnuLanguage.Enabled = Enable
'
'EHelp:
''If Not FromFullscreen Then mnuHelpContents.Enabled = Enable
''If Not FromFullscreen Then mnuAbout.Enabled = Enable
'
'EMacros:
'If EnableMacros <> 3 Then
'    mnuMacroCreate.Enabled = EnableMacros
'    mnuMacroLoad.Enabled = EnableMacros
'    For Z = mnuMacroRun.LBound To mnuMacroRun.UBound
'        mnuMacroRun(Z).Enabled = EnableMacros
'    Next
'Else
'    mnuMacroCreate.Enabled = Enable
'    mnuMacroLoad.Enabled = Enable
'    For Z = mnuMacroRun.LBound To mnuMacroRun.UBound
'        mnuMacroRun(Z).Enabled = Enable
'    Next
'End If

'
'If EnableMacros <> 3 Then mnuMacros.Enabled = EnableMacros Else mnuMacros.Enabled = Enable
End Sub

Public Sub EnableMenus(ByVal NewState As MenuStates, Optional ByVal ShouldTouchToolbars As Boolean = True)
Dim Z As Long, Enable As Boolean
Enable = NewState = mnsStandard

If ShouldTouchToolbars Then
    Select Case NewState
    Case mnsStandard
        MenuBar(1).Enabled = True
        MenuBar(2).Enabled = True
    Case mnsMacroGivens, mnsMacroResults, mnsMacroRun, mnsObjectSelection, mnsDemo
        MenuBar(1).Enabled = False
        MenuBar(2).Enabled = True
    'Case mnsFigureCreate, mnsFullScreen
    Case Else
        MenuBar(1).Enabled = False
        MenuBar(2).Enabled = False
    End Select
End If

'=========================================================
If mnuMacroResults.Visible Then mnuMacroResults.Visible = False
If mnuMacroSave.Visible Then mnuMacroSave.Visible = False
'=========================================================

EFile:
mnuNew.Enabled = Enable
mnuOpen.Enabled = Enable
mnuSave.Enabled = Enable
mnuSaveAs.Enabled = Enable
mnuExport.Enabled = Enable

mnuPrint.Enabled = Enable And Printers.Count > 0

For Z = 1 To mnuMRUFile.UBound
    mnuMRUFile(Z).Enabled = Enable
Next

'=========================================================

EEdit:
mnuUndo.Enabled = Enable
mnuRedo.Enabled = Enable
If Enable Then
    mnuUndo.Enabled = ActivityCount > 0
    mnuRedo.Enabled = UndoneActivityCount > 0
End If

mnuInsertLabel.Enabled = Enable
mnuInsertButton.Enabled = Enable
mnuCalculator.Enabled = Enable
mnuClearAll.Enabled = Enable
mnuFileProps.Enabled = Enable

'=========================================================

EView:
mnuShowAxes.Enabled = Enable
mnuShowGrid.Enabled = Enable
mnuShowRulers.Enabled = Enable

mnuShowStatusbar.Enabled = NewState <> mnsFullScreen

mnuShowToolbar.Enabled = Enable
mnuShowMainbar.Enabled = Enable

mnuPointList.Enabled = Enable
mnuFigureList.Enabled = Enable

mnuFullscreen.Enabled = (NewState = mnsStandard) Or (NewState = mnsFullScreen)

mnuDemo.Enabled = Enable
mnuDemoOptions.Enabled = Enable

'=========================================================

EFigures:
mnuFigCircles.Enabled = Enable
mnuFigConstruction.Enabled = Enable
mnuFigLines.Enabled = Enable
mnuFigMeasure.Enabled = Enable
mnuFigPoints.Enabled = Enable
mnuAnalytic.Enabled = Enable

'=========================================================

EOptions:
mnuSettings.Enabled = Enable
mnuLanguage.Enabled = Enable

'=========================================================

EHelp:

'=========================================================

EMacros:
Select Case NewState
Case mnsStandard
    mnuMacroCreate.Enabled = True
    mnuMacroLoad.Enabled = True
    For Z = mnuMacroRun.LBound To mnuMacroRun.UBound
        mnuMacroRun(Z).Enabled = True
    Next
    mnuMacroOrganize.Enabled = True
Case mnsMacroGivens
    mnuMacroCreate.Enabled = False
    mnuMacroLoad.Enabled = False
    For Z = mnuMacroRun.LBound To mnuMacroRun.UBound
        mnuMacroRun(Z).Enabled = False
    Next
    mnuMacroResults.Visible = True
    mnuMacroOrganize.Enabled = False
Case mnsMacroResults
    mnuMacroCreate.Enabled = False
    mnuMacroLoad.Enabled = False
    For Z = mnuMacroRun.LBound To mnuMacroRun.UBound
        mnuMacroRun(Z).Enabled = False
    Next
    mnuMacroResults.Visible = False
    mnuMacroSave.Visible = True
    mnuMacroOrganize.Enabled = False
Case mnsMacroRun
    mnuMacroCreate.Enabled = False
    mnuMacroLoad.Enabled = False
    For Z = mnuMacroRun.LBound To mnuMacroRun.UBound
        mnuMacroRun(Z).Enabled = False
    Next
    mnuMacroOrganize.Enabled = False
Case Else
'Case mnsDemo, mnsFigureCreate, mnsFullScreen, mnsObjectSelection
    mnuMacroCreate.Enabled = False
    mnuMacroLoad.Enabled = False
    For Z = mnuMacroRun.LBound To mnuMacroRun.UBound
        mnuMacroRun(Z).Enabled = False
    Next
    mnuMacroOrganize.Enabled = False
End Select

End Sub
