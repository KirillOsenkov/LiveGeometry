Imports GuiLabs.Canvas.Events
Imports GuiLabs.Canvas.Shapes

Public Interface IDGDocument
	Inherits IDrawable
	Inherits IMouseHandler

	Event NeedRedraw()
	Sub Resize(ByVal x As Integer, ByVal y As Integer)

	ReadOnly Property CmdFactory() As ICommandFactory
	'WriteOnly Property ViewControl() As System.Windows.Forms.Control

	'ReadOnly Property DocumentEvents() As IUIPrescriptionEvents
	'ReadOnly Property Notify() As IDocumentEventNotifications

End Interface
