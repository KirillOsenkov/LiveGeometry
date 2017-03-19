Imports DynamicGeometry

Public Class DGAppCommand

	Inherits Command

	Private mParentForm As frmParent
	Public Property ParentForm() As frmParent
		Get
			Return mParentForm
		End Get
		Set(ByVal Value As frmParent)
			mParentForm = Value
		End Set
	End Property

End Class
