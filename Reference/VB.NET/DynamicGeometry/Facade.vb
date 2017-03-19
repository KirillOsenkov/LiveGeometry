Public Class Facade

	Public Shared Function CreateDGDocument() As IDGDocument
		Return New DGDocument()
	End Function

    Public Shared Function CreateDGView() As DGView
        Return New DGView()
    End Function

End Class