Public Class CommandEditUndo
	Inherits DGCommand

	Public Overrides Sub OnClick()
		Me.ParentDocument.ActionManager.Undo()
        Me.ParentDocument.RaiseNeedRedraw()
	End Sub

	Public Overrides Property Enabled() As Boolean
		Get
			Return Me.ParentDocument.ActionManager.CanUndo
		End Get
		Set(ByVal Value As Boolean)
			'Throw New Exception("Cannot modify Enabled property of CommandEditUndo Command.")
		End Set
	End Property

End Class
