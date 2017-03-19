Namespace Actions

	' Represents the process of adding an IFigure object to the DGDocument.
	' Is bidirectional: Execute adds the figure and UnExecute removes it

	Friend Class AddFigureAction

		Inherits DGAction

		'====================================================================
		' Overrides for ExecuteCore & UnExecuteCore
		'====================================================================

		Protected Overrides Sub ExecuteCore()
			Doc.Figures.AddFigure(Figure)
			'Doc.View.RequestRedraw()
		End Sub

		Protected Overrides Sub UnExecuteCore()
			Doc.Figures.RemoveFigure(Figure)
			'Doc.View.RequestRedraw()
		End Sub

		'====================================================================
		' Constructing 
		'====================================================================

		Private Sub New()
		End Sub

		' returns a new AddFigureAction object, that can add/remove an IFigure to/from a DGDocument.
		' IFigure must be completely ready to be inserted into DGDocument:
		' all parents must be already pre-initialized with Figures from inside DGDocument.
		Public Shared Function Create(ByVal ExistingDocument As DGDocument, ByVal ExistingFigure As IFigure) As IAction
			Dim Act As AddFigureAction = New AddFigureAction()
			Act.Figure = ExistingFigure
			Act.Doc = ExistingDocument
			Return Act
		End Function

		'====================================================================
		' Private action data
		'====================================================================

		Private mFigure As IFigure
		Private Property Figure() As IFigure
			Get
				Return mFigure
			End Get
			Set(ByVal Value As IFigure)
				mFigure = Value
			End Set
		End Property

	End Class

End Namespace