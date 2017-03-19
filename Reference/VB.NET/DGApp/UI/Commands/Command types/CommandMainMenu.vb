Public Class CommandMainMenu
	Inherits DGAppCommand

	Public Sub New(ByVal Visible As Boolean, Optional ByVal Enabled As Boolean = True)
		Me.Visible = Visible
		Me.Enabled = Enabled
	End Sub

	Public Sub New()

	End Sub
End Class
