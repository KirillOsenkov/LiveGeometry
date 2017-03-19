Public Class UIPrescriptionEvents

	Implements IUIPrescriptionEvents

	Public Event ShouldDisplayStatus(ByVal Text As String) Implements DynamicGeometry.IUIPrescriptionEvents.ShouldDisplayStatus
	Public Event ShouldEnableUI(ByVal Enable As Boolean) Implements DynamicGeometry.IUIPrescriptionEvents.ShouldEnableUI
	Public Event ShouldSwitchTool(ByVal NewTool As DynamicGeometry.GeometryTool) Implements DynamicGeometry.IUIPrescriptionEvents.ShouldSwitchTool

#Region " Interface OUT "
	'=================================================================================================================================
	Friend Sub EnableUI(ByVal Enable As Boolean)
		RaiseEvent ShouldEnableUI(Enable)
	End Sub

	Friend Sub DisplayStatus(ByVal Text As String)
		RaiseEvent ShouldDisplayStatus(Text)
	End Sub

	Friend Sub SwitchTool(ByVal NewTool As GeometryTool)
		RaiseEvent ShouldSwitchTool(NewTool)
	End Sub
	'=================================================================================================================================
#End Region ' UI notifications ("UI prescriptions" - what should UI do)

End Class
