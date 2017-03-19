VERSION 5.00
Begin VB.Form frmToolTip 
   Appearance      =   0  'Flat
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   300
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   1440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.2
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Text"
      ForeColor       =   &H80000017&
      Height          =   228
      Left            =   24
      TabIndex        =   0
      Top             =   24
      Width           =   348
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartTime As Long

Private Sub Form_Unload(Cancel As Integer)
ToolTipShown = False
CurrentToolTipText = ""
End Sub

Private Sub tmrTimer1_Timer()
If Timer > StartTime + Val(tmrTimer1.Tag) Or Timer < 10 Then tmrTimer1.Enabled = False: Unload Me
End Sub
