Imports System.Collections.Generic

Public Class DGCommandFactory

	Inherits CommandFactory

	Friend Sub New(ByVal ExistingDocument As DGDocument)
		mParentDocument = ExistingDocument
	End Sub

	Private mParentDocument As DGDocument = Nothing	' Doc
	Friend ReadOnly Property ParentDocument() As DGDocument
		Get
			Return mParentDocument
		End Get
	End Property

	Protected Overrides Function CreateNewCommand(ByVal CommandName As String) As Command
		Dim NewCommand As DGCommand = Nothing

		Select Case CommandName
			Case CommandStrings.Edit.Undo
				NewCommand = New CommandEditUndo()

			Case CommandStrings.Edit.Redo
				NewCommand = New CommandEditRedo()

			Case CommandStrings.Mode.Drag
				NewCommand = New CommandModeDrag()
				NewCommand.Tooltip = "Move"

			Case CommandStrings.Tool.Point
				NewCommand = New CommandToolPoint()
				NewCommand.Tooltip = "Create a point"

			Case CommandStrings.Tool.Segment
				NewCommand = New CommandToolSegment()
				NewCommand.Tooltip = "Draw a segment"

			Case CommandStrings.Tool.Ray
				NewCommand = New CommandToolRay()
				NewCommand.Tooltip = "Draw a ray starting in a point and passing through another point"

			Case CommandStrings.Tool.Line
				NewCommand = New CommandToolLine()
				NewCommand.Tooltip = "Draw a line through two points"

			Case CommandStrings.Tool.MidPoint
				NewCommand = New CommandToolMidPoint()
				NewCommand.Tooltip = "Construct a midpoint between two points"

			Case CommandStrings.Tool.Circle
				NewCommand = New CommandToolCircle()
				NewCommand.Tooltip = "Draw a circle by center and a point on it"

			Case CommandStrings.Tool.CircleByRadius
				NewCommand = New CommandToolCircleByRadius()
				NewCommand.Tooltip = "Draw a circle by center and radius"

		End Select

		NewCommand.ParentDocument = ParentDocument

		Return NewCommand
	End Function

End Class
