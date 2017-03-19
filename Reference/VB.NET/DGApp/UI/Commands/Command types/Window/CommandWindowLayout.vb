Public Class CommandWindowLayout

	Inherits DGAppCommand

	Private NewLayout As System.Windows.Forms.MdiLayout

	Public Sub New(ByVal Layout As System.Windows.Forms.MdiLayout)
		NewLayout = Layout
		If Layout = MdiLayout.ArrangeIcons Then
			Me.Visible = False
		End If
	End Sub

	Public Overrides Sub OnClick()
		ParentForm.LayoutMdi(NewLayout)
	End Sub

End Class
