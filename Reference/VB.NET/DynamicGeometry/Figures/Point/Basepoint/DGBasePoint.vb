<FigureType( _
FigureTypeStrings.BasePoint, _
FigureCategoryStrings.Point, _
New String() {} _
)> _
Friend Class DGBasePoint

	Inherits DGPoint

	Public Sub New(ByVal ContainerDoc As DGDocument, ByVal X As Integer, ByVal Y As Integer)
		MyBase.New(ContainerDoc, CType(ContainerDoc.Paper.ActiveCoordinateSystem, CoordinateSystem))
		'Me.Style = AppearanceFactory.Instance.CreatePointAppearance(System.Drawing.Color.Black, System.Drawing.Color.LightGoldenrodYellow, PointShapeType.Circle, 3)
		Me.SetPhysicalCoordinates(X, Y)
	End Sub

	'Private Shared mFigureType As IFigureType
	'Public Shared Function GetFigureType() As IFigureType
	'	If mFigureType Is Nothing Then
	'		mFigureType = New FigureType(FigureTypeStrings.BasePoint, FigureCategoryStrings.Point)
	'	End If
	'	Return mFigureType
	'End Function

	'   Public Overrides ReadOnly Property FigureType() As DynamicGeometry.IFigureType
	'       Get
	'           Return GetFigureType()
	'       End Get
	'   End Property

End Class
