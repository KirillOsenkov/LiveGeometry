Imports System.Windows.Forms
Imports GuiLabs.Canvas.Events
Imports DynamicGeometry.Actions

Friend Class Dragger

	Inherits InteractiveBehaviour

	Public Sub New(ByVal ParentDoc As DGDocument)
		Doc = ParentDoc
		Reset()
	End Sub

	Private Dragging As Boolean = False
	Private P1 As IDGPoint
	Private ox As Integer, oy As Integer

	Private Transaction As MovePointAction

	Private dependents As IFigureList = Nothing

	Public Overrides Sub Reset()
		Dragging = False
		P1 = Nothing
		dependents = Nothing
	End Sub

	Public Sub MovePoint(ByVal Point1 As IDGPoint, ByVal x As Integer, ByVal y As Integer)
	End Sub

#Region " Overrides MouseEvents "
	'================================================================================
	Public Overrides Sub OnMouseDown(ByVal e As MouseEventArgsWithKeys)
		Dim i As IDGObject

		i = Doc.Figures.ObjectUnderPoint(CInt(e.X), CInt(e.Y))

		Dragging = False
		If Not i Is Nothing AndAlso TypeOf (i) Is IDGPoint Then
			Dragging = True
			P1 = CType(i, DGPoint)
			ox = e.X - P1.Coordinates.X
			oy = e.Y - P1.Coordinates.Y
			Transaction = MovePointAction.Create(Doc, P1)
			Doc.ActionManager.RecordAction(Transaction)

			Dim nav As New FigureListNavigator
			dependents = nav.GetAllDependentsSorted(P1)
		End If

	End Sub

	Public Overrides Sub OnMouseMove(ByVal e As MouseEventArgsWithKeys)
		If Dragging Then
			P1.MoveTo(CInt(e.X) - ox, CInt(e.Y) - oy)
			If dependents IsNot Nothing Then
				dependents.Recalculate()
			End If
			Doc.RaiseNeedRedraw()
		End If

		Dim x As Double = e.X, y As Double = e.Y
		' Doc.Paper.ActiveCoordinateSystem.ToLogical(x, y)
		' Doc.Request.DisplayStatus(x & "; " & y)

	End Sub

	Public Overrides Sub OnMouseUp(ByVal e As MouseEventArgsWithKeys)
		'If dependents IsNot Nothing Then

		'	Dim s As New Text.StringBuilder

		'	For Each f As IFigure In dependents
		'		s.AppendLine(DirectCast(f, Object).ToString())
		'	Next
		'	MessageBox.Show(s.ToString)
		'End If

		Reset()
	End Sub
	'================================================================================
#End Region

End Class