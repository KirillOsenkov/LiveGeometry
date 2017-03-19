Imports System.Collections.Generic

Friend Interface IFigureType

	Property Name() As String
	Property Category() As String
	Property Parents() As IList(Of String)

End Interface