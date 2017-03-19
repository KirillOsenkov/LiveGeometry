Imports GuiLabs.Canvas.Events
Imports GuiLabs.Canvas.Renderer

Friend MustInherit Class FigureCreator
	Inherits ObjectCreator

	Public Sub New(ByVal ParentDoc As DGDocument)
		Doc = ParentDoc
		Reset()
		InitializeFigureType()
	End Sub

	Protected CurrentParent As Integer = 0
	Protected NewType As IFigureType
	Protected NewParents As FigureList = New FigureList()
	Protected Transaction As Actions.MultiAction = New Actions.MultiAction()
	Private Temporary As IFigure
	Protected TempPoint As IDGPoint = Nothing

	Public MustOverride Function CreateFigure() As IFigure

	Public MustOverride Sub InitializeFigureType()

	'====================================================================

	Public Overrides Sub Reset()
		CurrentParent = 0
		NewParents.Clear()
		Transaction = New Actions.MultiAction()
		Temporary = Nothing
		TempPoint = Nothing
	End Sub

	Public Overrides Sub Finish()
		Dim NewFigure As IFigure = CreateFigure()
		Dim Action As Actions.IAction = Actions.AddFigureAction.Create(Doc, NewFigure)
		Transaction.InnerActions.Add(Action)
		Me.Doc.ActionManager.RecordAction(Transaction)
		MyBase.Finish()
	End Sub

	'====================================================================

	Private Function CreateNewPoint(ByVal x As Integer, ByVal y As Integer, Optional ByVal ShouldAddAction As Boolean = True) As IDGPoint
		Dim NewPoint As IDGPoint = New DGBasePoint(Doc, x, y)
		If ShouldAddAction Then
			Dim Action As Actions.IAction = Actions.AddFigureAction.Create(Doc, NewPoint)
			Transaction.InnerActions.Add(Action)
		End If
		Return NewPoint
	End Function

	Public Overrides Sub OnMouseDown(ByVal e As MouseEventArgsWithKeys)
		If e.IsRightButtonPressed Then
			AbortAndSetDefaultTool()
			Return
		End If

		Dim ObjectList As IFigureList = Doc.Figures.ObjectsUnderPoint(CInt(e.X), CInt(e.Y))
		Dim FigureToAdd As IFigure = ObjectList.FindFirstFigure(NewType.Parents(CurrentParent))

		If FigureToAdd Is Nothing And NewType.Parents(CurrentParent) = FigureCategoryStrings.Point Then
			FigureToAdd = CreateNewPoint(CInt(e.X), CInt(e.Y))
		End If

		If Not FigureToAdd Is Nothing Then
			NewParents.Add(FigureToAdd)
			Doc.RaiseNeedRedraw()
			CurrentParent += 1
		End If

		'====================================================================

		If Not TempPoint Is Nothing Then
			NewParents.Remove(TempPoint)
			TempPoint = Nothing
		End If

		If CurrentParent < NewType.Parents.Count Then
			If NewType.Parents(CurrentParent) = FigureCategoryStrings.Point Then
				TempPoint = CreateNewPoint(CInt(e.X), CInt(e.Y), False)
				TempPoint.Style = AppearanceFactory.Instance.FindPointAppearance(AppearanceFactory.Points.SelectedPoint)
				NewParents.Add(TempPoint)
				If CurrentParent = NewType.Parents.Count - 1 Then
					Temporary = CreateFigure()
				End If
				Doc.RaiseNeedRedraw()
			End If
		ElseIf CurrentParent = NewType.Parents.Count Then
			Temporary = Nothing
			Finish()
		End If

	End Sub

	Public Overrides Sub OnMouseMove(ByVal e As MouseEventArgsWithKeys)
		If Not TempPoint Is Nothing Then
			TempPoint.MoveTo(CInt(e.X), CInt(e.Y))
			If Temporary IsNot Nothing Then
				Temporary.Recalculate()
			End If
			Doc.RaiseNeedRedraw()
		End If
	End Sub

	Public Overrides Sub Draw(ByVal Renderer As IRenderer)
		MyBase.Draw(Renderer)

		If Not Temporary Is Nothing Then
			DirectCast(Temporary, IDGObject).Draw(Renderer)
		End If

		Dim i As IFigure
		For Each i In NewParents
			If i.FigureType.Category = FigureCategoryStrings.Point Then
				DirectCast(i, IDGPoint).Draw(Renderer)
			End If
		Next
	End Sub
End Class
