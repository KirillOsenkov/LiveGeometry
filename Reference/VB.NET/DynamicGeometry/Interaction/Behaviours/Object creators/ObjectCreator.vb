Friend MustInherit Class ObjectCreator

	Inherits InteractiveBehaviour

	' Clears all internal data; makes no changes to the outermost world.
	Public Overrides Sub Reset()
	End Sub

	' Aborts currently running operation.
	' Sets state Initial, calls Reset() and updates the screen.
	Public Overridable Sub Abort()
		mIsInitial = True
		Me.Reset()
        Me.Doc.RaiseNeedRedraw()
	End Sub

	' Process successful end of the current operation
	Public Overridable Sub Finish()
		Abort()
	End Sub

	Protected mIsInitial As Boolean = True
	Public Property IsInitial() As Boolean
		Get
			Return mIsInitial
		End Get
		Set(ByVal Value As Boolean)
			mIsInitial = Value
		End Set
	End Property

	Public Overridable Sub AbortAndSetDefaultTool()
		If Not IsInitial Then
			Abort()
		Else
			Me.Doc.CmdFactory.GetCommand(CommandStrings.Mode.Drag).OnClick()
		End If
	End Sub
End Class
