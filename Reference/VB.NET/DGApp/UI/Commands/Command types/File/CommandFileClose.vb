Public Class CommandFileClose

	Inherits DGAppCommand

	Public Overrides Sub OnClick()
		Dim m As frmChild = ParentForm.ActiveMdiChild
		If Not IsNothing(m) Then m.Close()
	End Sub

End Class
