Imports GuiLabs.Canvas
Imports GuiLabs.Canvas.Utils

Module DGMath

	Public Function Sqr(ByVal num As Double) As Double
		Return num * num
	End Function

    Public Function Distance(ByVal P1 As IMathPoint, ByVal P2 As IMathPoint) As Double
        Return Math.Sqrt((P2.x - P1.x) ^ 2 + (P2.y - P1.y) ^ 2)
    End Function

    Public Function Distance(ByVal P1 As CartesianPoint, ByVal P2 As CartesianPoint) As Double
        Return Math.Sqrt((P2.UnitsX - P1.UnitsX) ^ 2 + (P2.UnitsY - P1.UnitsY) ^ 2)
	End Function

	Public Function Distance(ByVal p1 As MathPoint, ByVal p2 As MathPoint) As Double
		Return Math.Sqrt(Sqr(p2.x - p1.x) + Sqr(p2.y - p1.y))
	End Function

	Public Function PixelDistance(ByVal P1 As Point, ByVal P2 As Point) As Integer
		Return CInt(Math.Sqrt((P2.X - P1.X) ^ 2 + (P2.Y - P1.Y) ^ 2))
	End Function

	Public Function GetPerpPoint(ByVal source As MathPoint, ByVal line As MathTwoPoints) As MathPoint
		Dim result As MathPoint

		If line.p1.y = line.p2.y Then
			result.x = source.x
			result.y = line.p1.y
		ElseIf line.p1.x = line.p2.x Then
			result.x = line.p1.x
			result.y = source.y
		Else
			Dim A As Double = (source - line.p1).SumOfSquares
			Dim B As Double = (source - line.p2).SumOfSquares
			Dim C As Double = (line.p1 - line.p2).SumOfSquares

			If C <> 0 Then
				Dim m As Double = (A + C - B) / (2 * C)
				result = ScalePointBetweenTwo(line.p1, line.p2, m)
			Else
				result = line.p1
			End If
		End If

		Return result
	End Function

	Public Function GetLineFromSegment(ByVal p1 As MathPoint, ByVal p2 As MathPoint, ByVal borders As MathTwoPoints) As MathTwoPoints
		Dim p3 As MathPoint, p4 As MathPoint

		If p1.x = p2.x And p1.y = p2.y Then
			p3 = p1
			p4 = p1
		ElseIf p1.x = p2.x Then
			p3.x = p1.x
			p4.x = p1.x
			p3.y = borders.p1.y
			p4.y = borders.p2.y
			If p1.y < p2.y Then Common.Swap(Of Double)(p3.y, p4.y)
		ElseIf p1.y = p2.y Then
			p3.y = p1.y
			p4.y = p1.y
			p3.x = borders.p1.x
			p4.x = borders.p2.x
			If p1.x < p2.x Then Common.Swap(Of Double)(p3.x, p4.x)
		Else
			Dim deltax As Double = p2.x - p1.x
			Dim deltay As Double = p2.y - p1.y
			Dim deltaxyRatio As Double = deltax / deltay
			Dim deltayxRatio As Double = deltay / deltax

			If deltay < 0 Then p3.y = borders.p1.y Else p3.y = borders.p2.y
			p3.x = p1.x + (p3.y - p1.y) * deltaxyRatio
			If p3.x < borders.p1.x Then
				p3.x = borders.p1.x
				p3.y = p1.y + (p3.x - p1.x) * deltayxRatio
			ElseIf p3.x > borders.p2.x Then
				p3.x = borders.p2.x
				p3.y = p1.y + (p3.x - p1.x) * deltayxRatio
			End If

			If deltax < 0 Then p4.x = borders.p2.x Else p4.x = borders.p1.x
			p4.y = p2.y + (p4.x - p2.x) * deltayxRatio
			If p4.y < borders.p1.y Then
				p4.y = borders.p1.y
				p4.x = p2.x + (p4.y - p2.y) * deltaxyRatio
			ElseIf p4.y > borders.p2.y Then
				p4.y = borders.p2.y
				p4.x = p2.x + (p4.y - p2.y) * deltaxyRatio
			End If
		End If

		GetLineFromSegment.p1 = p3
		GetLineFromSegment.p2 = p4
	End Function

	Public Function ScalePointBetweenTwo(ByVal source As MathPoint, ByVal dest As MathPoint, ByVal lambda As Double) As MathPoint
		Return New MathPoint(source.x + (dest.x - source.x) * lambda, source.y + (dest.y - source.y) * lambda)
	End Function

	Public Function ScaleNumberBetweenTwo(ByVal source As Double, ByVal dest As Double, ByVal lambda As Double) As Double
		Return source + (dest - source) * lambda
	End Function

	'Public Function GetPerpPoint(ByVal X As Double, ByVal Y As Double, ByVal XM As Double, ByVal YM As Double, ByVal XN As Double, ByVal YN As Double) As OnePoint
	'	Dim A As Double, B As Double, C As Double, m As Double
	'	Dim P As MathPoint

	'	If YN = YM Then
	'		P.X = X
	'		P.Y = YM
	'	ElseIf XM = XN Then
	'		P.X = XM
	'		P.Y = Y
	'	Else
	'		A = (X - XM) * (X - XM) + (Y - YM) * (Y - YM)
	'		B = (X - XN) * (X - XN) + (Y - YN) * (Y - YN)
	'		C = (XM - XN) * (XM - XN) + (YM - YN) * (YM - YN)
	'		If C <> 0 Then
	'			m = (A + C - B) / (2 * C)
	'			P.X = XM + (XN - XM) * m
	'			P.Y = YM + (YN - YM) * m
	'		Else
	'			P.X = XM
	'			P.Y = YM
	'		End If
	'	End If

	'	Return P
	'End Function

End Module
