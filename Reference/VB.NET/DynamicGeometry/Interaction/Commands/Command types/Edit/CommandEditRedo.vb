Public Class CommandEditRedo
	Inherits DGCommand

	Public Overrides Sub OnClick()
		Me.ParentDocument.ActionManager.Redo()
        Me.ParentDocument.RaiseNeedRedraw()
	End Sub

	Public Overrides Property Enabled() As Boolean
		Get
			Return Me.ParentDocument.ActionManager.CanRedo
		End Get
		Set(ByVal Value As Boolean)
			'Throw New Exception("Cannot modify Enabled property of CommandEditRedo Command.")
		End Set
	End Property

End Class
