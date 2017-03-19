Public Class CommandFileExit

	Inherits DGAppCommand

	Public Overrides Sub OnClick()
		ParentForm.Close()
	End Sub

End Class
