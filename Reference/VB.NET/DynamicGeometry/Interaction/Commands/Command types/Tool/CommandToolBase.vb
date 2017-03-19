Public MustInherit Class CommandToolBase
	Inherits DGCommand

	Public NotOverridable Overrides Sub OnClick()
		MyBase.OnClick()
		If Me.Name <> Me.ParentDocument.ActiveTool Then
			Me.ParentDocument.CmdFactory.GetCommand(Me.ParentDocument.ActiveTool).Checked = False
			Me.ParentDocument.ActiveTool = Name
			Me.ParentDocument.CmdFactory.GetCommand(Me.ParentDocument.ActiveTool).Checked = True
			OnClickCore()
		End If
	End Sub

	Protected MustOverride Sub OnClickCore()

End Class
