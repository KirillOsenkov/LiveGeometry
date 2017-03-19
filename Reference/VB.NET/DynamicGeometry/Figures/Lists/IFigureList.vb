Imports System.Collections.Generic

' An abstract list of IFigure objects
Friend Interface IFigureList

	Inherits IList(Of IFigure)

	Function FindFirstFigure(ByVal Category As String) As IFigure
	Sub Recalculate()

End Interface
