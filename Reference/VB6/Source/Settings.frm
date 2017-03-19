VERSION 5.00
Begin VB.Form frmSettings 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2
   Icon            =   "Settings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   4
      Left            =   4200
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   3
      Left            =   4200
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1695
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   2
      Left            =   4200
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1695
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   1
      Left            =   4200
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   63
      Top             =   840
      Width           =   1695
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   4170
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4935
      Index           =   3
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkToolSelectOnce 
         Caption         =   "&Tool select once"
         Height          =   255
         Left            =   240
         TabIndex        =   119
         Top             =   3000
         Width           =   3615
      End
      Begin VB.CheckBox chkShowStatusbar 
         Caption         =   "Show statusbar"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   2280
         Width           =   3615
      End
      Begin VB.CheckBox chkShowToolbar 
         Caption         =   "Show geometry toolbar"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   2040
         Width           =   3615
      End
      Begin VB.CheckBox chkShowMainToolbar 
         Caption         =   "Show main toolbar"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1800
         Width           =   3615
      End
      Begin VB.CheckBox chkShowTooltips 
         Caption         =   "Tooltips"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   3240
         Width           =   3615
      End
      Begin VB.CheckBox chkShowClock 
         Caption         =   "Clock"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   4200
         Width           =   3615
      End
      Begin VB.CheckBox chkShowCoords 
         Caption         =   "Show coords"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   3720
         Width           =   3615
      End
      Begin VB.CheckBox chkLoadCursors 
         Caption         =   "Load cursors"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   3480
         Width           =   3615
      End
      Begin VB.CheckBox chkShowRulers 
         Caption         =   "Show rulers"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Frame fraLanguage 
         Caption         =   "Language"
         Height          =   855
         Left            =   240
         TabIndex        =   89
         Top             =   240
         Width           =   3645
         Begin VB.OptionButton optLanguage 
            Caption         =   "Ukrainian"
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   40
            Tag             =   "2001"
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "German"
            Height          =   240
            Index           =   2
            Left            =   1800
            TabIndex        =   41
            Tag             =   "2000"
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "&English"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Tag             =   "0"
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optLanguage 
            Caption         =   "&Russian"
            Height          =   240
            Index           =   1
            Left            =   1800
            TabIndex        =   42
            Tag             =   "1"
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkGradient 
         Caption         =   "&Gradient fill objects"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   3960
         Width           =   3615
      End
      Begin DG.ctlColorBox csbToolbarColor 
         Height          =   252
         Left            =   240
         TabIndex        =   44
         Top             =   1440
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin DG.ctlColorBox csbRulerColor 
         Height          =   252
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   397
      End
      Begin VB.Label lblRulersColor 
         AutoSize        =   -1  'True
         Caption         =   "Ruler color"
         Height          =   216
         Left            =   600
         TabIndex        =   91
         Top             =   1200
         Width           =   3180
      End
      Begin VB.Label lblToolbarColor 
         AutoSize        =   -1  'True
         Caption         =   "Toolbar color"
         Height          =   216
         Left            =   600
         TabIndex        =   90
         Top             =   1440
         Width           =   3180
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4935
      Index           =   4
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Frame fraMacroCreate 
         Caption         =   "When creating a macro"
         Height          =   1215
         Left            =   240
         TabIndex        =   112
         Top             =   2880
         Width           =   3615
         Begin VB.CheckBox chkShowMacroResultsDialog 
            Caption         =   "Show ""Select results"" dialog"
            Height          =   375
            Left            =   120
            TabIndex        =   114
            Top             =   720
            Width           =   3375
         End
         Begin VB.CheckBox chkShowMacroCreateDialog 
            Caption         =   "Show ""Begin"" dialog"
            Height          =   495
            Left            =   120
            TabIndex        =   113
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame fraPrecision 
         Caption         =   "Decimal precision"
         Height          =   1575
         Left            =   240
         TabIndex        =   92
         Top             =   240
         Width           =   3615
         Begin VB.TextBox txtDistancePrecision 
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Text            =   "2"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtAnglePrecision 
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Text            =   "1"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtNumberPrecision 
            Height          =   315
            Left            =   120
            TabIndex        =   54
            Text            =   "2"
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblDigitNumber 
            Caption         =   "(number of digits after decimal)"
            Height          =   255
            Left            =   120
            TabIndex        =   118
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblDistancePrecision 
            Caption         =   "Distance decimal places"
            Height          =   315
            Left            =   840
            TabIndex        =   95
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label lblAnglePrecision 
            Caption         =   "Angle decimal places"
            Height          =   315
            Left            =   840
            TabIndex        =   94
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label lblNumberPrecision 
            Caption         =   "Number decimal places"
            Height          =   315
            Left            =   840
            TabIndex        =   93
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.TextBox txtSensitivity 
         Height          =   315
         Left            =   240
         TabIndex        =   57
         Text            =   "3"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtMacroAutoloadPath 
         Height          =   315
         Left            =   240
         TabIndex        =   59
         Top             =   4440
         Width           =   3255
      End
      Begin VB.CommandButton cmdMacroAutoloadPath 
         Caption         =   "..."
         Height          =   315
         Left            =   3480
         TabIndex        =   60
         Top             =   4440
         Width           =   375
      End
      Begin VB.CheckBox chkTransparentMetafile 
         Caption         =   "Export Transparent metafiles"
         Height          =   252
         Left            =   240
         TabIndex        =   58
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label lblSensitivity 
         Caption         =   "Cursor sensitivity (pixels):"
         Height          =   255
         Left            =   840
         TabIndex        =   97
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label lblAutoloadMacros 
         Caption         =   "Autoload-at-startup macro folder"
         Height          =   252
         Left            =   240
         TabIndex        =   96
         Top             =   4200
         Width           =   3612
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4935
      Index           =   0
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   120
      Width           =   4095
      Begin VB.VScrollBar vsbPointSize 
         Height          =   315
         Left            =   2640
         Max             =   30
         Min             =   2
         TabIndex        =   14
         Top             =   3120
         Value           =   4
         Width           =   240
      End
      Begin VB.CheckBox chkPointFill 
         Caption         =   "Fill"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox txtPointSize 
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   2280
         TabIndex        =   13
         Text            =   "12"
         Top             =   3120
         Width           =   360
      End
      Begin VB.Frame fraBasePoint 
         Caption         =   "Base point"
         Height          =   855
         Left            =   240
         TabIndex        =   107
         Top             =   120
         Width           =   3615
         Begin VB.ComboBox cmbPointShape 
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Settings.frx":0442
            Left            =   2040
            List            =   "Settings.frx":044C
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   600
         End
         Begin DG.ctlColorBox csbBasepointFillColor 
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin DG.ctlColorBox csbBasepointColor 
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.Label lblPointColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   210
            Left            =   480
            TabIndex        =   110
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblPointShape 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shape"
            Height          =   210
            Left            =   2760
            TabIndex        =   109
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblPointFillColor 
            Caption         =   "Fill color"
            Height          =   255
            Left            =   480
            TabIndex        =   108
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame fraPointNames 
         Caption         =   "Point names"
         Height          =   1095
         Left            =   240
         TabIndex        =   106
         Top             =   3600
         Width           =   3615
         Begin VB.CommandButton cmdPointLabelFont 
            Caption         =   "Name font"
            Height          =   375
            Left            =   1920
            TabIndex        =   17
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkAutoShowPointNames 
            Caption         =   "Auto show point names"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3375
         End
         Begin DG.ctlColorBox csbPointNameColor 
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   720
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.Label lblNameColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name color"
            Height          =   210
            Left            =   480
            TabIndex        =   16
            Top             =   720
            Width           =   810
         End
      End
      Begin VB.Frame fraFigurePoint 
         Caption         =   "Point on figure"
         Height          =   855
         Left            =   240
         TabIndex        =   102
         Top             =   1080
         Width           =   3615
         Begin VB.ComboBox cmbFigurePointShape 
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Settings.frx":0456
            Left            =   2040
            List            =   "Settings.frx":0460
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   600
         End
         Begin DG.ctlColorBox csbFigurepointFillColor 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin DG.ctlColorBox csbFigurepointColor 
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.Label lblFigurePointFillColor 
            Caption         =   "Fill color"
            Height          =   255
            Left            =   480
            TabIndex        =   105
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblFigurePointShape 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shape"
            Height          =   210
            Left            =   2760
            TabIndex        =   104
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblFigurePointColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   210
            Left            =   480
            TabIndex        =   103
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraDependentPoint 
         Caption         =   "Dependent point"
         Height          =   855
         Left            =   240
         TabIndex        =   98
         Top             =   2040
         Width           =   3615
         Begin VB.ComboBox cmbDependentPointShape 
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "Settings.frx":046A
            Left            =   2040
            List            =   "Settings.frx":0474
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   600
         End
         Begin DG.ctlColorBox csbDependentPointFillColor 
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin DG.ctlColorBox csbDependentPointColor 
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.Label lblDependentFillColor 
            Caption         =   "Fill color"
            Height          =   255
            Left            =   480
            TabIndex        =   101
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblDependentPointShape 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shape"
            Height          =   210
            Left            =   2760
            TabIndex        =   100
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblDependentColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   210
            Left            =   480
            TabIndex        =   99
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Label lblPointSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   210
         Left            =   3000
         TabIndex        =   111
         Top             =   3120
         Width           =   735
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4935
      Index           =   1
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Frame fraDefLocusProperties 
         Caption         =   "Locus properties"
         Height          =   2535
         Left            =   240
         TabIndex        =   77
         Top             =   2040
         Width           =   3615
         Begin VB.TextBox txtLocusDetailsHigh 
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Text            =   "640"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txtLocusDetails 
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Text            =   "64"
            Top             =   1680
            Width           =   495
         End
         Begin VB.ComboBox cmbLocusType 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "Settings.frx":047E
            Left            =   120
            List            =   "Settings.frx":0488
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   960
            Width           =   855
         End
         Begin VB.VScrollBar vsbLocusDrawWidth 
            Height          =   315
            Left            =   2160
            Max             =   16
            Min             =   1
            TabIndex        =   26
            Top             =   360
            Value           =   1
            Width           =   240
         End
         Begin VB.TextBox txtLocusWidth 
            ForeColor       =   &H80000012&
            Height          =   315
            Left            =   1800
            TabIndex        =   25
            Text            =   "1"
            Top             =   360
            Width           =   375
         End
         Begin DG.ctlColorBox csbLocusColor 
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.Label lblLocusColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   210
            Left            =   480
            TabIndex        =   82
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblLocusWidth 
            BackStyle       =   0  'Transparent
            Caption         =   "Draw width"
            Height          =   450
            Left            =   2520
            TabIndex        =   81
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLocusType 
            Caption         =   "Locus type"
            Height          =   255
            Left            =   1080
            TabIndex        =   80
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblLocusDetails 
            Caption         =   "Locus details"
            Height          =   375
            Left            =   720
            TabIndex        =   79
            Top             =   1680
            Width           =   2775
         End
         Begin VB.Label lblLocusDetailsHigh 
            Caption         =   "High quality locus details"
            Height          =   375
            Left            =   720
            TabIndex        =   78
            Top             =   2040
            Width           =   2775
         End
      End
      Begin VB.Frame fraDefLabelProps 
         Caption         =   "Label properties"
         Height          =   795
         Left            =   240
         TabIndex        =   75
         Top             =   1200
         Width           =   3615
         Begin VB.CommandButton cmdLabelFont 
            Caption         =   "Font"
            Height          =   375
            Left            =   1800
            TabIndex        =   23
            Top             =   240
            Width           =   1695
         End
         Begin DG.ctlColorBox csbLabelColor 
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.Label lblForeColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   255
            Left            =   480
            TabIndex        =   76
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame fraDefFigureProps 
         Caption         =   "Figure properties"
         Height          =   1035
         Left            =   240
         TabIndex        =   71
         Top             =   120
         Width           =   3615
         Begin VB.TextBox txtFigureDrawWidth 
            Height          =   315
            Left            =   1800
            TabIndex        =   20
            Text            =   "1"
            Top             =   360
            Width           =   375
         End
         Begin VB.VScrollBar vsbFigureDrawWidth 
            Height          =   315
            Left            =   2160
            Max             =   16
            Min             =   1
            TabIndex        =   21
            Top             =   360
            Value           =   1
            Width           =   240
         End
         Begin DG.ctlColorBox csbFigureFillColor 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin DG.ctlColorBox csbFigureColor 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.Label lblFigureColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   210
            Left            =   480
            TabIndex        =   74
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblFigureDrawWidth 
            Caption         =   "Draw width"
            Height          =   375
            Left            =   2520
            TabIndex        =   73
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblFigureFillColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fill color"
            Height          =   210
            Left            =   480
            TabIndex        =   72
            Top             =   600
            Width           =   1185
         End
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4935
      Index           =   2
      Left            =   120
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdLoadDefaultSettings 
         Caption         =   "Load Default settings"
         Height          =   495
         Left            =   240
         TabIndex        =   117
         Top             =   4320
         Width           =   3615
      End
      Begin VB.CommandButton cmdSaveAsDefault 
         Caption         =   "Save these settings as Default"
         Height          =   495
         Left            =   240
         TabIndex        =   116
         Top             =   3720
         Width           =   3615
      End
      Begin VB.CheckBox chkShowAxesMarks 
         Caption         =   "Show axes marks"
         Height          =   255
         Left            =   240
         TabIndex        =   115
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtWallpaper 
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Top             =   3000
         Width           =   3255
      End
      Begin VB.CommandButton cmdWallpaper 
         Caption         =   "..."
         Height          =   315
         Left            =   3480
         TabIndex        =   38
         Top             =   3000
         Width           =   375
      End
      Begin VB.Frame fraInterfaceColors 
         Caption         =   "Colors"
         Height          =   1335
         Left            =   240
         TabIndex        =   83
         Top             =   1320
         Width           =   3615
         Begin DG.ctlColorBox csbGridColor 
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin DG.ctlColorBox csbAxesColor 
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin DG.ctlColorBox csbPaperColor2 
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin DG.ctlColorBox csbPaperColor1 
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.Label lblAxesColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Axes color"
            Height          =   210
            Left            =   480
            TabIndex        =   87
            Top             =   720
            Width           =   1635
         End
         Begin VB.Label lblGridColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grid color"
            Height          =   210
            Left            =   480
            TabIndex        =   86
            Top             =   960
            Width           =   1635
         End
         Begin VB.Label lblPaperColor2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Paper color 2"
            Height          =   210
            Left            =   480
            TabIndex        =   85
            Top             =   480
            Width           =   1635
         End
         Begin VB.Label lblPaperColor1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Paper color 1"
            Height          =   210
            Left            =   480
            TabIndex        =   84
            Top             =   240
            Width           =   2115
         End
      End
      Begin VB.CheckBox chkShowGrid 
         Caption         =   "Show grid"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   3615
      End
      Begin VB.CheckBox chkShowAxes 
         Caption         =   "Show axes"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   3615
      End
      Begin VB.CheckBox chkPaperGradient 
         Caption         =   "&Gradient fill paper"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label lblWallpaper 
         Caption         =   "Background picture"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   2760
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ActiveTab As Long
Dim TabCaptions() As String

Dim OX As Long, OY As Long
Dim unlCancel As Boolean

Dim OldSetLanguage As Long
Dim OldGroupTools As Boolean
Dim OldPaperGradient As Boolean
Dim OldPrecisionCheckSum As Long

Dim OldPointFontName As String
Dim OldPointFontSize As Long
Dim OldPointFontBold As Boolean
Dim OldPointFontItalic As Boolean
Dim OldPointFontUnderline As Boolean
Dim OldPointFontCharset As Long

Dim OldFontName As String
Dim OldFontSize As Long
Dim OldFontBold As Boolean
Dim OldFontItalic As Boolean
Dim OldFontUnderline As Boolean
Dim OldFontCharset As Long

Private Sub chkPointFill_Click()
csbBasepointFillColor.Enabled = -chkPointFill.Value
csbFigurepointFillColor.Enabled = csbBasepointFillColor.Enabled
csbDependentPointFillColor.Enabled = csbBasepointFillColor.Enabled
End Sub

Private Sub cmdCancel_Click()
Cancel
End Sub

Private Sub cmdLoadDefaultSettings_Click()
csbPaperColor1.Color = setcolPaper1
csbPaperColor2.Color = setcolPaper2
csbAxesColor.Color = setcolAxes
csbGridColor.Color = setcolGrid
chkShowAxes.Value = -setShowAxes
chkShowGrid.Value = -setShowGrid
chkPaperGradient.Value = -setPaperGradient
End Sub

Private Sub cmdOK_Click()
OK
End Sub

Private Sub cmdSaveAsDefault_Click()
setcolPaper1 = csbPaperColor1.Color
setcolPaper2 = csbPaperColor2.Color
setcolGrid = csbGridColor.Color
setcolAxes = csbAxesColor.Color
setPaperGradient = -chkPaperGradient.Value
setShowAxes = -chkShowAxes.Value
setShowGrid = -chkShowGrid.Value

SaveSetting AppName, "Interface", "PaperGradientFill", Format(-CInt(setPaperGradient))
SaveSetting AppName, "Interface", "ShowAxes", Format(-CInt(setShowAxes))
SaveSetting AppName, "Interface", "ShowGrid", Format(-CInt(setShowGrid))
SaveSetting AppName, "Interface", "PaperColor1", setcolPaper1
SaveSetting AppName, "Interface", "PaperColor2", setcolPaper2
SaveSetting AppName, "Interface", "GridColor", setcolGrid
SaveSetting AppName, "Interface", "AxesColor", setcolAxes
End Sub

Private Sub optLanguage_Click(Index As Integer)
setLanguage = optLanguage(Index).Tag
FillSetStrings
PrepareContainers
End Sub

Private Sub chkGradient_Click()
SetGrad
End Sub

Private Sub cmdLabelFont_Click()
CD.FontName = OldFontName
CD.FontSize = OldFontSize
CD.FontBold = OldFontBold
CD.FontItalic = OldFontItalic
CD.FontUnderline = OldFontUnderline
CD.FontCharset = OldFontCharset
CD.Color = 0
CD.ShowFont
If CD.Cancelled Then Exit Sub
OldFontName = CD.FontName
OldFontSize = CD.FontSize
OldFontBold = CD.FontBold
OldFontItalic = CD.FontItalic
OldFontUnderline = CD.FontUnderline
OldFontCharset = CD.FontCharset
If CD.Color <> 0 Then csbLabelColor.Color = CD.Color
End Sub

Private Sub cmdPointLabelFont_Click()
CD.FontName = OldPointFontName
CD.FontSize = OldPointFontSize
CD.FontBold = OldPointFontBold
CD.FontItalic = OldPointFontItalic
CD.FontUnderline = OldPointFontUnderline
CD.FontCharset = OldPointFontCharset
CD.ShowFont
If CD.Cancelled Then Exit Sub
OldPointFontName = CD.FontName
OldPointFontSize = CD.FontSize
OldPointFontBold = CD.FontBold
OldPointFontItalic = CD.FontItalic
OldPointFontUnderline = CD.FontUnderline
OldPointFontCharset = CD.FontCharset
End Sub

Private Sub cmdMacroAutoloadPath_Click()
Dim strPath As String
strPath = BrowseForFolder(GetString(ResLocateMacroAutoloadPath))
If strPath = "" Then txtMacroAutoloadPath.Text = "": Exit Sub
If Dir(strPath, vbDirectory Or vbHidden Or vbReadOnly Or vbSystem) <> "" Then txtMacroAutoloadPath.Text = strPath
End Sub

Private Sub cmdWallpaper_Click()
Dim S As String
CD.Filter = "*.BMP; *.JPG; *.GIF; *.WMF; *.RLE; *.ICO|*.BMP;*.JPG;*.GIF;*.WMF;*.RLE;*.ICO"
CD.FileName = ""
CD.Flags = &H1000 + &H4
If txtWallpaper.Text <> "" And Dir(txtWallpaper.Text) <> "" Then CD.InitDir = RetrieveDir(txtWallpaper.Text) Else CD.InitDir = LastTexturePath
If Not IsValidPath(CD.InitDir) Then CD.InitDir = ProgramPath
CD.DialogTitle = GetString(ResOpen)
CD.ShowOpen
If CD.Cancelled Then Exit Sub
If Dir(CD.FileName) = "" Then
    txtWallpaper.Text = ""
    S = RetrieveDir(CD.FileName)
    If IsValidPath(S) Then LastTexturePath = AddDirSep(S)
Else
    txtWallpaper.Text = CD.FileName
End If
End Sub

'=======================================
'                           Form Events
'=======================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
'    Case vbKeyEscape
'        unlCancel = True
'        Unload Me
'    Case vbKeyReturn
'        unlCancel = False
'        Unload Me
    Case vbKeyTab
        If Shift And vbCtrlMask Then
            If Shift And vbShiftMask Then
                ActivateTab IIf(ActiveTab = 0, picTab.UBound, ActiveTab - 1)
            Else
                ActivateTab IIf(ActiveTab = picTab.UBound, picTab.LBound, ActiveTab + 1)
            End If
        End If
End Select
End Sub

Private Sub Form_Load()
FirstInit
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Outit
End Sub

'==========================================================================
'==========================================================================
'==========================================================================


Private Sub picTab_Click(Index As Integer)
If Index <> ActiveTab Then ActivateTab Index
End Sub

Private Sub FillSetStrings()
Dim Z As Long

Caption = GetString(ResOptions)
cmdCancel.Caption = GetString(ResCancel)

'PrepareContainers
ReDim TabCaptions(picTab.LBound To picTab.UBound)

For Z = picTab.LBound To picTab.UBound
    TabCaptions(Z) = GetString(ResSetPoints + Z * 2)
Next

'===========================================
' Frame captions
'===========================================

'fraPoints.Caption = GetString(ResSetNewPointProperties)
'fraOtherElements.Caption = GetString(ResSetFigures)
'fraPaperProps.Caption = GetString(ResSetPaper)
'fraInterface.Caption = GetString(ResSetInterface)
'fraMisc.Caption = GetString(ResSetMisc)

fraBasePoint.Caption = GetString(ResSetBasePoints)
fraFigurePoint.Caption = GetString(ResSetFigurePoints)
fraDependentPoint.Caption = GetString(ResSetDependentPoints)
fraPointNames.Caption = GetString(ResSetPointNames)

fraPrecision.Caption = GetString(ResDecimalPrecision)
fraDefFigureProps.Caption = GetString(ResMnuFigureProperties)
fraDefLabelProps.Caption = GetString(ResMnuLabelProperties)
fraDefLocusProperties.Caption = GetString(ResLocusProps)

'===========================================
' Point properties
'===========================================

chkPointFill.Caption = GetString(ResFill)
lblPointSize.Caption = GetString(ResSize)
lblNameColor.Caption = GetString(ResNameColor)

lblPointColor.Caption = GetString(ResColor)
lblFigurePointColor.Caption = GetString(ResColor)
lblDependentColor.Caption = GetString(ResColor)
lblPointFillColor.Caption = GetString(ResFill)
lblFigurePointFillColor.Caption = GetString(ResFill)
lblDependentFillColor.Caption = GetString(ResFill)
lblPointShape.Caption = GetString(ResShape)
lblFigurePointShape.Caption = GetString(ResShape)
lblDependentPointShape.Caption = GetString(ResShape)

'===========================================
' Figure, label and loci properties
'===========================================

lblFigureColor.Caption = GetString(ResColor)
lblFigureFillColor.Caption = GetString(ResFill)
lblFigureDrawWidth.Caption = GetString(ResDrawWidth)
lblLocusWidth.Caption = GetString(ResDrawWidth)
lblForeColor.Caption = GetString(ResColor)
lblLocusColor.Caption = GetString(ResLocusColor)
lblLocusDetails = GetString(ResLocusDetails)
lblLocusDetailsHigh = GetString(ResLocusDetailsHigh)
'fraInterface.Caption = GetString(ResSetInterface)
fraInterfaceColors.Caption = GetString(ResColors)
cmdPointLabelFont.Caption = GetString(ResFont)
cmdLabelFont.Caption = GetString(ResFont)
lblLocusType.Caption = GetString(ResLocusType)

'===========================================
' Paper properties
'===========================================

lblPaperColor1.Caption = GetString(ResPaperColor) & " 1"
lblPaperColor2.Caption = GetString(ResPaperColor) & " 2"
lblGridColor.Caption = GetString(ResGridColor)
lblAxesColor.Caption = GetString(ResAxesColor)

chkShowAxes.Caption = GetString(ResShowAxes)
chkShowGrid.Caption = GetString(ResShowGrid)
chkPaperGradient.Caption = "&" & GetString(ResSetPaperGradient)

lblWallpaper.Caption = GetString(ResWallpaper)

'===========================================
' Interface properties
'===========================================

lblRulersColor.Caption = GetString(ResRulerColor)
lblToolbarColor.Caption = GetString(ResToolbarColor)

chkGradient.Caption = "&" & GetString(ResSetGradientFill)
chkLoadCursors.Caption = GetString(ResLoadCursors)

chkShowAxesMarks.Caption = GetString(ResShowAxesMarks)
chkShowRulers.Caption = GetString(ResShowRulers)
chkShowMainToolbar.Caption = GetString(ResShowMainbar)
chkShowStatusbar.Caption = GetString(ResShowStatusbar)
chkShowToolbar.Caption = GetString(ResShowToolbar)

chkShowClock.Caption = GetString(ResClock)
chkShowTooltips.Caption = GetString(ResShowTooltips)
chkShowCoords.Caption = GetString(ResCursorCoordinates)

'===========================================
' Miscellaneous
'===========================================

chkTransparentMetafile.Caption = GetString(ResSaveTransparentEMF)
chkToolSelectOnce.Caption = "&" & GetString(ResSetSelectToolOnce)
chkAutoShowPointNames.Caption = GetString(ResAutoShowPointName)
lblAutoloadMacros.Caption = GetString(ResMacroAutoloadPath)

lblSensitivity.Caption = GetString(ResCursorSensitivity)
lblDistancePrecision = GetString(ResDistancePrecision)
lblAnglePrecision = GetString(ResAnglePrecision)
lblNumberPrecision = GetString(ResNumberPrecision)
lblDigitNumber = GetString(ResDigitNumber)

fraMacroCreate.Caption = GetString(ResMacroDuringCreation)
chkShowMacroCreateDialog.Caption = GetString(ResMacroShowCreateDialog)
chkShowMacroResultsDialog.Caption = GetString(ResMacroShowResultsDialog)

cmdSaveAsDefault.Caption = GetString(ResSaveAsDefaults)
cmdLoadDefaultSettings.Caption = GetString(ResLoadDefaults)
End Sub

Private Sub SetGrad()
Dim OldSetGradientFill As Boolean, Z As Long

If chkGradient.Value Then
    'Gradient hDC, colStatusGradient, BackColor, 0, 0, ScaleWidth, ScaleHeight, False
    Cls
    OldSetGradientFill = setGradientFill
    setGradientFill = True
    ShadowControl picContainer(0)
    For Z = picTab.LBound To picTab.UBound
        ShadowControl picTab(Z)
    Next
    ShadowControl cmdCancel
    ShadowControl cmdOK
    setGradientFill = OldSetGradientFill
    
    Refresh
Else
    Cls
End If

PrepareContainers
End Sub

Private Sub PrepareContainers()
Dim R1 As RECT, TL As Long, Z As Long
Const TabDist = 4

'ReDim TabCaptions(picTab.LBound To picTab.UBound)

For Z = picTab.LBound To picTab.UBound
    'TabCaptions(Z) = GetString(ResSetPoints + Z * 2)
    TL = picContainer(ActiveTab).Left + picContainer(ActiveTab).Width
    picTab(Z).Move IIf(Z = ActiveTab, TL - 2, TL), picContainer(ActiveTab).Top + Z * (picTab(0).Height + TabDist)
    R1.Right = picContainer(Z).ScaleWidth
    R1.Bottom = picContainer(Z).ScaleHeight
    DrawEdge picContainer(Z).hDC, R1, EDGE_RAISED, BF_RECT
    PaintTab Z
Next

End Sub

Private Sub ActivateTab(ByVal Index As Long)
Dim Z As Long, T As Long
'Dim DC As Long, SrcDC As Long, nWidth As Long, nHeight As Long, X As Long, Y As Long
'Dim lpPoint As POINTAPI
'Const SettingsAnimSteps = 100

picContainer(Index).Move picContainer(ActiveTab).Left, picContainer(ActiveTab).Top, picContainer(ActiveTab).Width, picContainer(ActiveTab).Height

'picContainer(Index).Visible = True
'picContainer(ActiveTab).Visible = False
'picContainer(ActiveTab).AutoRedraw = False
'picContainer(Index).AutoRedraw = False

'DC = GetWindowDC(GetDesktopWindow)
'nWidth = picContainer(ActiveTab).ScaleWidth
'nHeight = picContainer(ActiveTab).ScaleHeight
'SrcDC = picContainer(Index).hDC 'GetWindowDC(picContainer(Index).hWnd)
'ClientToScreen picContainer(ActiveTab).hWnd, lpPoint
'X = lpPoint.X
'Y = lpPoint.Y
'For Z = SettingsAnimSteps To 0 Step -1
'    BitBlt DC, X + nWidth * Z / SettingsAnimSteps, Y, nWidth - nWidth * Z / SettingsAnimSteps, nHeight, SrcDC, 0, 0, SRCCOPY
'    'DoEvents
'    For Q = 1 To 10000
'        W = E
'    Next
'Next

'picContainer(Index).AutoRedraw = True
'picContainer(ActiveTab).AutoRedraw = True
picContainer(Index).Visible = True
picContainer(ActiveTab).Visible = False

picTab(ActiveTab).Left = picTab(ActiveTab).Left + 2
picTab(Index).Left = picTab(Index).Left - 2
Me.Line (picTab(Index).Left, picTab(Index).Top)-(picTab(Index).Left + picTab(Index).Width + Shadow * 2, picTab(Index).Top + picTab(Index).Height + Shadow * 2), BackColor, BF
If chkGradient.Value Then
    ShadowControl picContainer(Index)
    For Z = picTab.LBound To picTab.UBound
        ShadowControl picTab(Z)
    Next
End If
T = ActiveTab
ActiveTab = Index
PaintTab T
PaintTab ActiveTab
End Sub

Private Sub PaintTab(ByVal Index As Long)
Dim R1 As RECT, DTOptions As Long
Dim S As Size

picTab(Index).Cls

If chkGradient.Value Then
    picTab(Index).BackColor = IIf(Index = ActiveTab, colActiveTab, DarkenColor(picContainer(Index).BackColor, 0.85))
Else
    picTab(Index).BackColor = picContainer(Index).BackColor 'IIf(Index = ActiveTab, picContainer(Index).BackColor, colInActiveTab)
End If

R1.Right = picTab(Index).ScaleWidth
R1.Bottom = picTab(Index).ScaleHeight
If chkGradient.Value Then Gradient picTab(Index).hDC, picContainer(Index).BackColor, picTab(Index).BackColor, 0, 0, R1.Right, R1.Bottom, True
DrawEdge picTab(Index).hDC, R1, EDGE_RAISED, BF_TOP Or BF_BOTTOM Or BF_RIGHT

'===================================================
' Now draw the text caption
'===================================================

'picTab(Index).CurrentX = (picTab(Index).ScaleWidth - picTab(Index).TextWidth(TabCaptions(Index))) / 2
'picTab(Index).CurrentY = (picTab(Index).ScaleHeight - picTab(Index).TextHeight(TabCaptions(Index))) / 2
'picTab(Index).Print TabCaptions(Index)

Const HorizMargin = 4 ' left and right border size (in pixels) to leave on a button

If picTab(Index).TextWidth(TabCaptions(Index)) <= R1.Right - 2 * HorizMargin Then
    DTOptions = DT_CENTER Or DT_VCENTER Or DT_WORDBREAK Or DT_SINGLELINE
    R1.Left = 0
    R1.Top = 0
    R1.Right = picTab(Index).ScaleWidth
    R1.Bottom = picTab(Index).ScaleHeight
Else
    ' Calculate the required size of the rectangle
    R1.Left = HorizMargin
    R1.Top = 0
    R1.Right = picTab(Index).ScaleWidth - HorizMargin
    R1.Bottom = R1.Top
    
    DTOptions = DT_WORDBREAK Or DT_CALCRECT
    DrawText picTab(Index).hDC, TabCaptions(Index), Len(TabCaptions(Index)), R1, DTOptions
    OffsetRect R1, 0, (picTab(Index).ScaleHeight - R1.Bottom + R1.Top) \ 2
    
    R1.Right = picTab(Index).ScaleWidth - HorizMargin
    DTOptions = DT_CENTER Or DT_VCENTER Or DT_WORDBREAK
End If

DrawText picTab(Index).hDC, TabCaptions(Index), Len(TabCaptions(Index)), R1, DTOptions
End Sub

'=======================================
'                       Validation routines
'=======================================

Private Sub txtAnglePrecision_Validate(Cancel As Boolean)
If Not IsNumeric(txtAnglePrecision.Text) Then
    MsgBox GetString(ResDecimalPrecision) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 0 " + GetString(ResMsgTo) & " " & MaxPrecision & ".", vbInformation
    Cancel = True
Else
    If Val(txtAnglePrecision.Text) < 0 Or Val(txtAnglePrecision.Text) > MaxPrecision Then
        MsgBox GetString(ResDecimalPrecision) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 0 " + GetString(ResMsgTo) & " " & MaxPrecision & ".", vbInformation
        Cancel = True
    Else
        setAnglePrecision = Int(Val(txtAnglePrecision.Text))
        setFormatAngle = "0" & IIf(setAnglePrecision > 0, ".", "") & String(setAnglePrecision, "0")
    End If
End If
End Sub

Private Sub txtDistancePrecision_Validate(Cancel As Boolean)
If Not IsNumeric(txtDistancePrecision.Text) Then
    MsgBox GetString(ResDecimalPrecision) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 0 " + GetString(ResMsgTo) & " " & MaxPrecision & ".", vbInformation
    Cancel = True
Else
    If Val(txtDistancePrecision.Text) < 0 Or Val(txtDistancePrecision.Text) > MaxPrecision Then
        MsgBox GetString(ResDecimalPrecision) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 0 " + GetString(ResMsgTo) & " " & MaxPrecision & ".", vbInformation
        Cancel = True
    Else
        setDistancePrecision = Int(Val(txtDistancePrecision.Text))
        setFormatDistance = "0" & IIf(setDistancePrecision > 0, ".", "") & String(setDistancePrecision, "0")
    End If
End If
End Sub

Private Sub txtNumberPrecision_Validate(Cancel As Boolean)
If Not IsNumeric(txtNumberPrecision.Text) Then
    MsgBox GetString(ResDecimalPrecision) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 0 " + GetString(ResMsgTo) & " " & MaxPrecision & ".", vbInformation
    Cancel = True
Else
    If Val(txtNumberPrecision.Text) < 0 Or Val(txtNumberPrecision.Text) > MaxPrecision Then
        MsgBox GetString(ResDecimalPrecision) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 0 " + GetString(ResMsgTo) & " " & MaxPrecision & ".", vbInformation
        Cancel = True
    Else
        setNumberPrecision = Int(Val(txtNumberPrecision.Text))
        setFormatNumber = "0" & IIf(setNumberPrecision > 0, ".", "") & String(setNumberPrecision, "0")
    End If
End If
End Sub

'==============================
'           Sensitivity validation
'==============================
Private Sub txtSensitivity_Validate(Cancel As Boolean)
If Not Visible Or Not IsNumeric(txtSensitivity.Text) Then MsgBox GetString(ResCursorSensitivity) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 1 " + GetString(ResMsgTo) + " 16.", vbInformation: Cancel = True: Exit Sub
If Val(txtSensitivity.Text) < 1 Or Val(txtSensitivity.Text) > 16 Then
    MsgBox GetString(ResCursorSensitivity) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 1 " + GetString(ResMsgTo) + " 16.", vbInformation
    Cancel = True
End If
End Sub

'==============================
'           Locus details validation
'==============================
Private Sub txtLocusDetails_Validate(Cancel As Boolean)
If Not IsNumeric(txtLocusDetails.Text) Then
    MsgBox GetString(ResLocusDetails) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " 10 " & GetString(ResMsgTo) & " " & MaxDynamicLocusDetails & ".", vbInformation
    Cancel = True
Else
    If Val(txtLocusDetails.Text) < 10 Or Val(txtLocusDetails.Text) > MaxDynamicLocusDetails Then
        MsgBox GetString(ResLocusDetails) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " 10 " & GetString(ResMsgTo) & " " & MaxDynamicLocusDetails & "."
        Cancel = True
    Else
        setLocusDetails = Val(txtLocusDetails.Text)
    End If
End If
End Sub

Private Sub txtLocusDetailsHigh_Validate(Cancel As Boolean)
If Not IsNumeric(txtLocusDetailsHigh.Text) Then
    MsgBox GetString(ResLocusDetailsHigh) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " 10 " & GetString(ResMsgTo) & " " & MaxDynamicLocusDetailsHigh & ".", vbInformation
    Cancel = True
Else
    If Val(txtLocusDetailsHigh.Text) < 0 Or Val(txtLocusDetailsHigh.Text) > 4095 Then
        MsgBox GetString(ResLocusDetailsHigh) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " 10 " & GetString(ResMsgTo) & " " & MaxDynamicLocusDetailsHigh & ".", vbInformation
        Cancel = True
    Else
        setLocusDetailsHigh = Val(txtLocusDetailsHigh.Text)
    End If
End If
End Sub

'==============================
'           Point size validation
'==============================
Private Sub txtPointSize_Change()
If Not Visible Or Not IsNumeric(txtPointSize.Text) Then Exit Sub
If Val(txtPointSize.Text) <> Int(Val(txtPointSize.Text)) Then Exit Sub
If txtPointSize < MinPointSize Or txtPointSize > MaxPointSize Then Exit Sub
If txtPointSize.Enabled Then
    vsbPointSize.Enabled = False
    vsbPointSize.Value = MaxPointSize + MinPointSize - Val(txtPointSize)
    vsbPointSize.Enabled = True
End If
End Sub

Private Sub txtPointSize_Validate(Cancel As Boolean)
If Not Visible Or Not IsNumeric(txtPointSize.Text) Then MsgBox GetString(ResSize) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " " & MinPointSize & " " & GetString(ResMsgTo) & " " & MaxPointSize & ".", vbInformation: Cancel = True: Exit Sub
If Val(txtPointSize.Text) < MinPointSize Or Val(txtPointSize.Text) > MaxPointSize Then
    MsgBox GetString(ResSize) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) & " " & MinPointSize & " " & GetString(ResMsgTo) & " " & MaxPointSize & ".", vbInformation
    Cancel = True
End If
End Sub

Private Sub vsbPointSize_Change()
If Not Visible Then Exit Sub
If vsbPointSize.Value < MinPointSize Or vsbPointSize.Value > MaxPointSize Then Exit Sub
txtPointSize.Enabled = False
txtPointSize = MaxPointSize + MinPointSize - vsbPointSize.Value
txtPointSize.Enabled = True
txtPointSize.SelStart = 0
txtPointSize.SelLength = Len(txtPointSize)
'txtPointSize.SetFocus
End Sub

'==============================
'           Figure drawwidth validation
'==============================
Private Sub txtFigureDrawWidth_Change()
If Not Visible Or Not IsNumeric(txtFigureDrawWidth.Text) Then Exit Sub
If Val(txtFigureDrawWidth.Text) <> Int(Val(txtFigureDrawWidth.Text)) Then Exit Sub
If txtFigureDrawWidth < 1 Or txtFigureDrawWidth > MaxDrawWidth Then Exit Sub
If txtFigureDrawWidth.Enabled Then
    vsbFigureDrawWidth.Enabled = False
    vsbFigureDrawWidth.Value = MaxDrawWidth + 1 - Val(txtFigureDrawWidth)
    vsbFigureDrawWidth.Enabled = True
End If
End Sub

Private Sub txtFigureDrawWidth_Validate(Cancel As Boolean)
If Not Visible Or Not IsNumeric(txtFigureDrawWidth.Text) Then MsgBox GetString(ResDrawWidth) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " 1 " & GetString(ResMsgTo) & " " & MaxDrawWidth & ".", vbInformation: Cancel = True: Exit Sub
If Val(txtFigureDrawWidth.Text) < 1 Or Val(txtFigureDrawWidth.Text) > MaxDrawWidth Then
    MsgBox GetString(ResDrawWidth) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 1 " + GetString(ResMsgTo) & " " & MaxDrawWidth & ".", vbInformation
    Cancel = True
End If
End Sub

Private Sub vsbFigureDrawWidth_Change()
If Not Visible Then Exit Sub
If vsbFigureDrawWidth.Value < 1 Or vsbFigureDrawWidth.Value > MaxDrawWidth Then Exit Sub
txtFigureDrawWidth.Enabled = False
txtFigureDrawWidth = MaxDrawWidth + 1 - vsbFigureDrawWidth.Value
txtFigureDrawWidth.Enabled = True
txtFigureDrawWidth.SelStart = 0
txtFigureDrawWidth.SelLength = Len(txtFigureDrawWidth)
'txtFigureDrawWidth.SetFocus
End Sub

'==============================
'           Locus drawwidth validation
'==============================
Private Sub txtLocusWidth_Change()
If Not Visible Or Not IsNumeric(txtLocusWidth.Text) Then Exit Sub
If Val(txtLocusWidth.Text) <> Int(Val(txtLocusWidth.Text)) Then Exit Sub
If txtLocusWidth < 1 Or txtLocusWidth > MaxDrawWidth Then Exit Sub
If txtLocusWidth.Enabled Then
    vsbLocusDrawWidth.Enabled = False
    vsbLocusDrawWidth.Value = MaxDrawWidth + 1 - Val(txtLocusWidth)
    vsbLocusDrawWidth.Enabled = True
End If
End Sub

Private Sub txtLocusWidth_Validate(Cancel As Boolean)
If Not Visible Or Not IsNumeric(txtLocusWidth.Text) Then MsgBox GetString(ResDrawWidth) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 1 " + GetString(ResMsgTo) & " " & MaxDrawWidth & ".", vbInformation: Cancel = True: Exit Sub
If Val(txtLocusWidth.Text) < 1 Or Val(txtLocusWidth.Text) > MaxDrawWidth Then
    MsgBox GetString(ResDrawWidth) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 1 " + GetString(ResMsgTo) & " " & MaxDrawWidth & ".", vbInformation
    Cancel = True
End If
End Sub

Private Sub vsbLocusDrawWidth_Change()
If Not Visible Then Exit Sub
If vsbLocusDrawWidth.Value < 1 Or vsbLocusDrawWidth.Value > MaxDrawWidth Then Exit Sub
txtLocusWidth.Enabled = False
txtLocusWidth = MaxDrawWidth + 1 - vsbLocusDrawWidth.Value
txtLocusWidth.Enabled = True
txtLocusWidth.SelStart = 0
txtLocusWidth.SelLength = Len(txtLocusWidth)
'txtLocusWidth.SetFocus
End Sub

'===================================
' Called when the form loads for the first time
'===================================
Private Sub FirstInit()
cmdOK.Top = 8 + picContainer(0).Height - cmdOK.Height
cmdCancel.Top = cmdOK.Top

FillSetStrings
End Sub

Private Sub Init()
' Called when the dialog is displayed
'===================================
InitPublicVariables

ActiveTab = 0
picContainer(0).Move 8, 8

FormMain.Enabled = False
unlCancel = False
'PrepareContainers

'===================================
' Reserve some of the old settings
'===================================

OldPointFontName = setdefPointFontName
OldPointFontSize = setdefPointFontSize
OldPointFontBold = setdefPointFontBold
OldPointFontItalic = setdefPointFontItalic
OldPointFontUnderline = setdefPointFontUnderline
OldPointFontCharset = setdefPointFontCharset

OldFontName = setdefLabelFont
OldFontSize = setdefLabelFontSize
OldFontBold = setdefLabelBold
OldFontItalic = setdefLabelItalic
OldFontUnderline = setdefLabelUnderline
OldFontCharset = setdefLabelCharset

OldPrecisionCheckSum = setNumberPrecision * 10000 + setDistancePrecision * 100 + setAnglePrecision
OldSetLanguage = setLanguage
OldGroupTools = setGroupTools
OldPaperGradient = nGradientPaper

'===================================
' Fill settings
'===================================

optLanguage(Switch(setLanguage = langEnglish, 0, setLanguage = langRussian, 1, setLanguage = langGerman, 2, setLanguage = langUkrainian, 3)).Value = True

chkGradient.Value = -setGradientFill
chkToolSelectOnce.Value = -setToolSelectOnce
chkPaperGradient.Value = -nGradientPaper
chkShowAxes.Value = -nShowAxes
chkShowGrid.Value = -nShowGrid
chkShowAxesMarks.Value = -setShowAxesMarks

chkShowRulers.Value = -setShowRulers
chkShowMainToolbar.Value = -setShowMainbar
chkShowToolbar.Value = -setShowToolbar
chkShowStatusbar.Value = -setShowStatusbar

chkShowCoords.Value = -setShowCoord
chkAutoShowPointNames.Value = -setAutoShowPointName
chkShowClock.Value = -setShowClock
chkShowTooltips.Value = -setShowTooltips
chkTransparentMetafile.Value = -setTransparentEMF
chkLoadCursors.Value = -setLoadCursors

txtNumberPrecision.Text = setNumberPrecision
txtDistancePrecision.Text = setDistancePrecision
txtAnglePrecision.Text = setAnglePrecision
txtSensitivity.Text = setCursorSensitivity
txtMacroAutoloadPath.Text = setMacroAutoloadPath
txtLocusDetails.Text = setLocusDetails
txtLocusDetailsHigh.Text = setLocusDetailsHigh
txtWallpaper.Text = setWallpaper

'===================================
' Fill color selection buttons
'===================================
csbRulerColor.Color = setcolRuler
csbToolbarColor.Color = setcolToolbar

csbBasepointColor.Color = setdefcolPoint
csbBasepointFillColor.Color = setdefcolPointFill
csbFigurepointColor.Color = setdefcolFigurePoint
csbFigurepointFillColor.Color = setdefcolFigurePointFill
csbDependentPointColor.Color = setdefcolDependentPoint
csbDependentPointFillColor.Color = setdefcolDependentPointFill
csbPointNameColor.Color = setdefcolPointName

csbFigureColor.Color = setdefcolFigure
csbFigureFillColor.Color = setdefcolFigureFill
csbLocusColor.Color = setdefcolLocus
csbLabelColor.Color = setdefcolLabelColor

csbPaperColor1.Color = nPaperColor1
csbPaperColor2.Color = nPaperColor2
csbGridColor.Color = nGridColor
csbAxesColor.Color = nAxesColor

'===================================
' Fill miscellaneous
'===================================

cmbPointShape.ListIndex = Sgn(setdefPointShape - 1)
cmbFigurePointShape.ListIndex = Sgn(setdefFigurePointShape - 1)
cmbDependentPointShape.ListIndex = Sgn(setdefDependentPointShape - 1)
cmbLocusType.ListIndex = setdefLocusType

chkShowMacroCreateDialog.Value = -setShowMacroCreateDialog
chkShowMacroResultsDialog.Value = -setShowMacroResultsDialog

txtPointSize.Text = setdefPointSize
txtFigureDrawWidth.Text = setdefFigureDrawWidth
txtLocusWidth.Text = setdefLocusDrawWidth
vsbPointSize.Value = MaxPointSize + MinPointSize - setdefPointSize
vsbFigureDrawWidth.Value = MaxDrawWidth + 1 - setdefFigureDrawWidth
vsbLocusDrawWidth.Value = MaxDrawWidth + 1 - setdefLocusDrawWidth

chkPointFill.Value = 1 - setdefPointFill
chkPointFill_Click

'===================================
' ... and finally show the dialog
'===================================
SetGrad

Visible = True
End Sub

Private Function Outit() As Boolean
' Take up actions to process settings dialog close events
'====================================================
On Local Error Resume Next

Dim OldSetShowMainbar As Boolean
Dim OldSetShowToolbar As Boolean
Dim OldSetShowStatusbar As Boolean
Dim OldSetShowRulers As Boolean

Dim OldSetGradientFill As Boolean
Dim OldRulerColor As Long, OldToolbarColor As Long
Dim OldFontTransparent As Boolean

Dim NewSens As Double, Z As Long
Dim ShouldResizeForm As Boolean

'===================================

FormMain.Enabled = True
FormMain.SetFocus
setLanguage = OldSetLanguage
If unlCancel Then Exit Function ' Don't touch anything if Cancelled

LockWindowUpdate FormMain.hWnd ' prevent window flickering

'===================================
' Reserve old settings for comparison later...
'===================================
OldSetShowMainbar = setShowMainbar
OldSetShowToolbar = setShowToolbar
OldSetShowStatusbar = setShowStatusbar
OldSetShowRulers = setShowRulers

OldRulerColor = setcolRuler
OldToolbarColor = setcolToolbar
OldSetGradientFill = setGradientFill

'===================================
' Restore current point font characteristics... Why do I need it?
'===================================

setdefPointFontName = OldPointFontName
setdefPointFontCharset = OldPointFontCharset
setdefPointFontSize = OldPointFontSize
setdefPointFontBold = OldPointFontBold
setdefPointFontItalic = OldPointFontItalic
setdefPointFontUnderline = OldPointFontUnderline

Paper.FontName = setdefPointFontName
Paper.FontSize = setdefPointFontSize
Paper.FontBold = setdefPointFontBold
Paper.FontItalic = setdefPointFontItalic
Paper.FontUnderline = setdefPointFontUnderline
Paper.Font.Charset = setdefPointFontCharset

'=================================
' Set new color settings
'=================================

setcolRuler = csbRulerColor.Color
setcolToolbar = csbToolbarColor.Color

setdefcolPoint = csbBasepointColor.Color
setdefcolPointFill = csbBasepointFillColor.Color
setdefcolFigurePoint = csbFigurepointColor.Color
setdefcolFigurePointFill = csbFigurepointFillColor.Color
setdefcolDependentPoint = csbDependentPointColor.Color
setdefcolDependentPointFill = csbDependentPointFillColor.Color
setdefcolPointName = csbPointNameColor.Color

setdefcolFigure = csbFigureColor.Color
setdefcolFigureFill = csbFigureFillColor.Color
setdefcolLocus = csbLocusColor.Color
setdefcolLabelColor = csbLabelColor.Color

'setcolPaper1 = csbPaperColor1.Color
'setcolPaper2 = csbPaperColor2.Color
'setcolGrid = csbGridColor.Color
'setcolAxes = csbAxesColor.Color
nPaperColor1 = csbPaperColor1.Color
nPaperColor2 = csbPaperColor2.Color
If nAxesColor <> csbAxesColor.Color Then
    nAxesColor = csbAxesColor.Color
    If WereActiveAxesAdded Then
        Figures(nActiveX).ForeColor = nAxesColor
        Figures(nActiveY).ForeColor = nAxesColor
    End If
End If
nAxesColor = csbAxesColor.Color
nGridColor = csbGridColor.Color

'===================================
' Display paper elements...
'===================================

'setPaperGradient = -chkPaperGradient.Value
'setShowAxes = -chkShowAxes.Value
'setShowGrid = -chkShowGrid.Value
nGradientPaper = -chkPaperGradient.Value
nShowAxes = -chkShowAxes.Value
If nShowAxes <> -chkShowAxes.Value Then
    ToggleAxes -chkShowAxes.Value
End If
nShowGrid = -chkShowGrid.Value

'===================================
' Other defaults...
'===================================

setdefPointShape = cmbPointShape.ListIndex * 2 + 1
setdefFigurePointShape = cmbFigurePointShape.ListIndex * 2 + 1
setdefDependentPointShape = cmbDependentPointShape.ListIndex * 2 + 1
setdefPointSize = Val(txtPointSize.Text)
setdefFigureDrawWidth = txtFigureDrawWidth.Text
setdefLocusDrawWidth = txtLocusWidth.Text
setdefLocusType = cmbLocusType.ListIndex
setdefPointFill = 1 - chkPointFill.Value

' Restore current label font characteristics...
setdefLabelFont = OldFontName
setdefLabelCharset = OldFontCharset
setdefLabelFontSize = OldFontSize
setdefLabelBold = OldFontBold
setdefLabelItalic = OldFontItalic
setdefLabelUnderline = OldFontUnderline
setdefLabelTransparent = OldFontTransparent

'====================================================
'                                   Macro autoload path
'====================================================
If Dir(txtMacroAutoloadPath.Text, vbDirectory Or vbHidden Or vbReadOnly Or vbSystem) <> "" Then setMacroAutoloadPath = txtMacroAutoloadPath.Text
If (GetFileAttributes(setMacroAutoloadPath) And vbDirectory) <> vbDirectory Then setMacroAutoloadPath = ""
'====================================================

'====================================================
'                                           Wallpaper
'====================================================
If setWallpaper <> txtWallpaper.Text Then
    If Dir(txtWallpaper.Text) = "" Then txtWallpaper.Text = ""
    setWallpaper = txtWallpaper.Text
    If setWallpaper <> "" Then
        LastTexturePath = RetrieveDir(setWallpaper)
        If IsValidPath(LastTexturePath) Then SaveSetting AppName, "Paths", "LastTexturePath", LastTexturePath
        PrepareWallPaper
    End If
    ShouldResizeForm = True
    SaveSetting AppName, "Interface", "Wallpaper", setWallpaper
End If

'====================================================
'                               Miscellaneous
'====================================================
setShowAxesMarks = -chkShowAxesMarks.Value
setGradientFill = -chkGradient.Value
setAutoShowPointName = -chkAutoShowPointNames.Value
setCursorSensitivity = Val(txtSensitivity.Text)
setShowTooltips = -chkShowTooltips.Value
setTransparentEMF = -chkTransparentMetafile.Value
setToolSelectOnce = -chkToolSelectOnce.Value
setLoadCursors = -chkLoadCursors.Value

setShowMacroCreateDialog = -chkShowMacroCreateDialog.Value
setShowMacroResultsDialog = -chkShowMacroResultsDialog.Value

For Z = optLanguage.LBound To optLanguage.UBound
    If optLanguage(Z).Value Then setLanguage = optLanguage(Z).Tag
Next Z

'===================================
' Display interface elements...
'===================================

setShowRulers = -chkShowRulers.Value
setShowMainbar = -chkShowMainToolbar.Value
setShowToolbar = -chkShowToolbar.Value
setShowStatusbar = -chkShowStatusbar.Value

setShowClock = -chkShowClock.Value
setShowCoord = -chkShowCoords.Value

If setGradientFill <> OldSetGradientFill Or OldSetShowRulers <> setShowRulers Or OldSetShowMainbar <> setShowMainbar Or OldSetShowToolbar <> setShowToolbar Or OldSetShowStatusbar <> setShowStatusbar Then
    ShouldResizeForm = True
    FormMain.PrepareFormControls True
Else
    If OldRulerColor <> setcolRuler Then FormMain.PrepareRulers True
    If OldToolbarColor <> setcolToolbar Then FormMain.PrepareToolbar
    FormMain.PrepareStatusBar True
End If

'===================================
' Sensitivity of cursor in pixels
'===================================
NewSens = setCursorSensitivity
ToLogicalLength NewSens
Sensitivity = NewSens


'===================================
' Save it all...
'===================================

SaveSettings

'===================================
' Refill strings with a new language, if necessary
'===================================

If OldSetLanguage <> setLanguage Then
    FormMain.FillStrings
    CreateMenus
End If

If ShouldResizeForm Then FormMain.Form_Resize

RecalcScrollAssociatedInfo
UpdateLabels

'=====================================

ActivateTab 0

DoEvents
LockWindowUpdate 0

'=====================================

PaperCls
ShowAll
End Function

Public Sub OK()
unlCancel = False
Outit
Me.Visible = False
End Sub

Public Sub Cancel()
unlCancel = True
Outit
Me.Visible = False
End Sub

Public Sub ShowSettings()
Init
Me.Visible = True
Me.Refresh
End Sub

Private Sub InitPublicVariables()
picContainer(0).Visible = True
picContainer(1).Visible = False
picContainer(2).Visible = False
picContainer(3).Visible = False
picContainer(4).Visible = False

ActiveTab = 0

OX = 0
OY = 0
unlCancel = False

OldSetLanguage = 0
OldGroupTools = False
OldPaperGradient = False
OldPrecisionCheckSum = 0

OldPointFontName = ""
OldPointFontSize = 0
OldPointFontBold = False
OldPointFontItalic = False
OldPointFontUnderline = False
OldPointFontCharset = 0

OldFontName = ""
OldFontSize = 0
OldFontBold = False
OldFontItalic = False
OldFontUnderline = False
OldFontCharset = 0
End Sub
