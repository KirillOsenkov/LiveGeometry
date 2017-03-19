Imports System.Windows.Forms
Imports System.Drawing
Imports GuiLabs.Canvas
Imports GuiLabs.Canvas.Renderer

Public Class DGView

    Inherits DrawWindow
    Implements IDGView

#Region "Document"

    Private WithEvents mDocument As IDGDocument
    Public Property Document() As IDGDocument
        Get
            Return mDocument
        End Get
        Set(ByVal value As IDGDocument)
            Me.DefaultMouseHandler = Nothing
            mDocument = value
            If mDocument IsNot Nothing Then
                Me.DefaultMouseHandler = mDocument
            End If
        End Set
    End Property

#End Region

#Region " Paint and resize "

    Public Sub OnRepaint(ByVal renderer As IRenderer) Handles Me.Repaint
        If Document IsNot Nothing Then
            Document.Draw(renderer)
        End If
    End Sub

    Public Sub DocumentNeedsRedraw() Handles mDocument.NeedRedraw
        Me.Redraw()
    End Sub

    Public Sub OnViewResize(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Resize
        If Document IsNot Nothing Then
            Document.Resize(Me.ClientSize.Width, Me.ClientSize.Height)
        End If
        Me.Redraw()
    End Sub

#End Region

End Class
