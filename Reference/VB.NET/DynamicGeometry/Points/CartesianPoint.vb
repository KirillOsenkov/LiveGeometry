Imports GuiLabs.Canvas

Friend Class CartesianPoint
	Inherits Point

	Private mUnitsX As Double
	Public Property UnitsX() As Double
		Get
			Return mUnitsX
		End Get
		Set(ByVal value As Double)
			mUnitsX = value
		End Set
	End Property

	Private mUnitsY As Double
	Public Property UnitsY() As Double
		Get
			Return mUnitsY
		End Get
		Set(ByVal value As Double)
			mUnitsY = value
		End Set
	End Property

	Public Function GetMathPoint() As MathPoint
		Return New MathPoint(UnitsX, UnitsY)
	End Function

	Public Sub SetLogical(ByVal newUnitsX As Double, ByVal newUnitsY As Double)
		UnitsX = newUnitsX
		UnitsY = newUnitsY
	End Sub

	Public Sub SetLogical(ByVal fromMathPoint As MathPoint)
		SetLogical(fromMathPoint.x, fromMathPoint.y)
	End Sub

	Public Function ToLogicalString() As String
		Return "(" + UnitsX.ToString() + "; " + UnitsY.ToString() + ")"
	End Function

End Class
