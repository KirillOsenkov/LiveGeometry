Imports System.Windows.Forms
Imports GuiLabs.Canvas.Events
Imports GuiLabs.Canvas.Renderer

Friend MustInherit Class InteractiveBehaviour

    Inherits MouseHandler
	Implements IBehaviour

	Public mDoc As DGDocument
	Public Property Doc() As DynamicGeometry.DGDocument Implements DynamicGeometry.IBehaviour.Doc
		Get
			Return mDoc
		End Get
		Set(ByVal Value As DynamicGeometry.DGDocument)
			mDoc = Value
		End Set
	End Property

	' Abort current operation, clear all data
	' and return to original state.

	' TODO: TOTHINK: Should this refresh screen or not? Who should refresh screen when aborted?
	' Clears all internal data; makes no changes to the outermost world.
    Public MustOverride Sub Reset() Implements IBehaviour.Reset

    Public Overridable Sub Draw(ByVal Renderer As IRenderer) Implements IBehaviour.Draw
        Doc.Figures.Draw(Renderer)
    End Sub

End Class
