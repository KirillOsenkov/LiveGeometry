Namespace Actions

	'====================================================================
	' The main class which is responsible for all transactions 
	' (modifications of DGDocument data structure)

	' "Facade" for the Actions namespace.
	'====================================================================

	Friend Class ActionCenter

		Public Sub New(ByVal ParentDoc As DGDocument)
			Doc = ParentDoc
		End Sub

		Private mDoc As DGDocument
		Public Property Doc() As DGDocument
			Get
				Return mDoc
			End Get
			Set(ByVal Value As DGDocument)
				mDoc = Value
			End Set
		End Property

		'====================================================================
		' 
		'====================================================================

		Public Sub RecordAction(ByVal ExistingAction As IAction)
			History.RecordAction(ExistingAction)
			History.MoveForward()
		End Sub

		'====================================================================
		' Undo & Redo
		'====================================================================

		Public Sub Undo()
			History.MoveBack()
		End Sub

		Public Sub Redo()
			History.MoveForward()
		End Sub

		Public Function CanUndo() As Boolean
			Return History.CanMoveBack
		End Function

		Public Function CanRedo() As Boolean
			Return History.CanMoveForward
		End Function

		'====================================================================
		' Internal data structure to keep track of all recorded actions.

		' Here implemented: SimpleHistory - a linear doubly linked-list of actions.
		'====================================================================

		Private mHistory As IActionHistory = New SimpleHistory()
		Protected Property History() As IActionHistory
			Get
				Return mHistory
			End Get
			Set(ByVal Value As IActionHistory)
				mHistory = Value
			End Set
		End Property

		'==============================================================================
		' Public Add* methods - adding new elements to the list
		'==============================================================================

		'Public Function AddNewPoint(ByVal x As Integer, ByVal y As Integer, ByVal FrameOfReference As ICoordinateSystem) As IDGPoint
		'	'Dim NewPoint As IDGPoint = Doc.Figures.Factory.CreateBasePoint(x, y, FrameOfReference)
		'	Me.AddFigure(NewPoint)
		'	Return NewPoint
		'End Function

		'Public Function AddNewSegment(ByVal EndPoint1 As IDGPoint, ByVal EndPoint2 As IDGPoint) As IDGLine
		'	Dim NewSegment As IDGLine = Doc.Figures.Factory.CreateSegment(EndPoint1, EndPoint2)
		'	Me.AddFigure(NewSegment)
		'	Return NewSegment
		'End Function

		'==============================================================================
		' Protected helper-methods
		'==============================================================================

		Public Sub AddFigure(ByVal ExistingFigure As IFigure)
			Dim Action As IAction = Actions.AddFigureAction.Create(Doc, ExistingFigure)
			RecordAction(Action)
		End Sub

	End Class
End Namespace