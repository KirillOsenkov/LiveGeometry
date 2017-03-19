Imports GuiLabs.Canvas.Renderer

Friend Class VisualObjectList
    Inherits GeometryObjectList

    Public Sub Draw(ByVal Renderer As IRenderer)
        Dim i As IDGObject
        For Each i In mList
			If i.FigureType.Category <> FigureCategoryStrings.Point Then
				i.Draw(Renderer)
			End If
		Next
        For Each i In mList
			If i.FigureType.Category = FigureCategoryStrings.Point Then
				i.Draw(Renderer)
			End If
		Next
    End Sub

    '====================================================================
    ' Input: Physical pixels x and y
    '====================================================================
    Public Function ObjectUnderPoint(ByVal x As Integer, ByVal y As Integer) As IDGObject
        Dim i As IDGObject
        For Each i In mList
            If i.IsPointOver(x, y) Then Return i
        Next
        Return Nothing
    End Function

    Public Function ObjectsUnderPoint(ByVal x As Integer, ByVal y As Integer) As IFigureList
        Dim i As IDGObject
        Dim NewList As IFigureList = New FigureList()
        For Each i In mList
            If i.IsPointOver(x, y) Then NewList.Add(i)
        Next
        Return NewList
    End Function

    'Public Function EnsureReturnPoint(ByVal x As Integer, ByVal y As Integer) As DGPoint
    '	Dim i As IDGObject = ObjectUnderPoint(x, y)
    '	' TODO: VisualObjectList::EnsureReturnPoint
    'End Function

    'Public Sub SetForeColorOfAllPoints(ByVal c As System.Drawing.Color)
    '	Dim i As DGObject
    '	For Each i In Me
    '		CType(i, DGBasePoint).Appearance.Pen.Color = c
    '	Next
    'End Sub

End Class
