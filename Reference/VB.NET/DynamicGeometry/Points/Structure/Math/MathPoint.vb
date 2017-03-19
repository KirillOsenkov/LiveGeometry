Friend Structure MathPoint

	Public x As Double
	Public y As Double

	Public Sub New(ByVal newx As Double, ByVal newy As Double)
		x = newx
		y = newy
	End Sub

#Region "Add"

	Public Sub Add(ByVal dx As Double, ByVal dy As Double)
		x += dx
		y += dy
	End Sub

	Public Sub Add(ByVal delta As Double)
		Add(delta, delta)
	End Sub

	Public Sub Add(ByVal p As MathPoint)
		Add(p.X, p.Y)
	End Sub

#End Region

#Region "Set"

	Public Sub SetCoords(ByVal x As Double, ByVal y As Double)
		x = x
		y = y
	End Sub

	Public Sub SetCoords(ByVal sizeOfBoth As Double)
		SetCoords(sizeOfBoth, sizeOfBoth)
	End Sub

	Public Sub SetCoords(ByVal p As MathPoint)
		SetCoords(p.x, p.y)
	End Sub

	Public Sub Set0()
		SetCoords(0, 0)
	End Sub

#End Region

	Public Shared Operator +(ByVal p1 As MathPoint, ByVal p2 As MathPoint) As MathPoint
		Return New MathPoint(p1.x - p2.x, p1.y - p2.y)
	End Operator

	Public Shared Operator -(ByVal p1 As MathPoint, ByVal p2 As MathPoint) As MathPoint
		Return New MathPoint(p1.x - p2.x, p1.y - p2.y)
	End Operator

	Public Shared Operator -(ByVal p1 As MathPoint) As MathPoint
		Return New MathPoint(-p1.x, -p1.y)
	End Operator

	Public Function SumOfSquares() As Double
		Return x * x + y * y
	End Function

End Structure

Friend Structure MathTwoPoints
	Public p1 As MathPoint
	Public p2 As MathPoint

	Public Sub New(ByVal point1 As MathPoint, ByVal point2 As MathPoint)
		p1 = point1
		p2 = point2
	End Sub
End Structure