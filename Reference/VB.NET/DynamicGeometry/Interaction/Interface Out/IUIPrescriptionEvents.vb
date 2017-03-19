Public Interface IUIPrescriptionEvents

	Event ShouldEnableUI(ByVal Enable As Boolean)
	Event ShouldDisplayStatus(ByVal Text As String)
	Event ShouldSwitchTool(ByVal NewTool As DynamicGeometry.GeometryTool)

End Interface
