Imports GuiLabs.Canvas

Namespace Actions

    Friend Class MovePointAction
        Inherits ToggleAction

        Protected Sub New()

        End Sub

		Private dependents As IFigureList = Nothing

        Public Shared Function Create(ByVal ExistingDocument As DGDocument, ByVal Point1 As IDGPoint) As MovePointAction
            Dim Act As MovePointAction = New MovePointAction()
            Act.Doc = ExistingDocument
            Act.Point = Point1
			Act.ExecuteCount = 1

			Dim nav As New FigureListNavigator()
			Act.dependents = nav.GetAllDependentsSorted(Point1)

            Act.OldPosition.Set(Point1.Coordinates)
            Return Act
        End Function

        Protected Overrides Sub ExecuteCore()
			Point.MoveTo(NewPosition)
			dependents.Recalculate()
            Doc.RaiseNeedRedraw()
        End Sub

        Protected Overrides Sub UnExecuteCore()
            NewPosition.Set(Point.Coordinates)
			Point.MoveTo(OldPosition)
			dependents.Recalculate()
        End Sub

        Private OldPosition As Point = New Point()
        Private NewPosition As Point = New Point()

        Private mPoint As IDGPoint
        Private Property Point() As IDGPoint
            Get
                Return mPoint
            End Get
            Set(ByVal Value As IDGPoint)
                mPoint = Value
            End Set
        End Property
    End Class

End Namespace
