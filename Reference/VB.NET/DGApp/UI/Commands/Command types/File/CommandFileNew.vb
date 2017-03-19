Public Class CommandFileNew

	Inherits DGAppCommand

	Public Overrides Sub OnClick()
		Dim NewDoc As New CDocument(ParentForm)
		ParentForm.Documents.Add(NewDoc)
		'ParentForm.CmdFactory.UpdateCommands()
	End Sub

End Class
