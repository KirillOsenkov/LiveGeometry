VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "About DG"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   HelpContextID   =   61
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WasResponce As Boolean

Private Sub Form_Deactivate()
WasResponce = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
WasResponce = True
End Sub

Private Sub Form_Load()
WasResponce = False
Me.Visible = True
Me.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
WasResponce = True
End Sub
