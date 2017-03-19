Public Class CommandStrings

	Public Const ToolPrefix As String = "Tool"

	Public Class Edit
		Public Const Undo As String = "Edit.Undo"
		Public Const Redo As String = "Edit.Redo"
	End Class

	Public Class Tool
		Public Const Point As String = ToolPrefix + ".Point"
		Public Const Segment As String = ToolPrefix + ".Segment"
		Public Const Ray As String = ToolPrefix + ".Ray"
		Public Const Line As String = ToolPrefix + ".Line"
		Public Const MidPoint As String = ToolPrefix + ".MidPoint"
		Public Const Circle As String = ToolPrefix + ".Circle"
		Public Const CircleByRadius As String = ToolPrefix + ".CircleByRadius"
	End Class

	Public Class Mode
		Public Const Drag As String = "Mode.Drag"
	End Class

End Class