Public Class frmChild

	Public Sub New()
		MyBase.New()

		' This call is required by the Windows Form Designer.
		InitializeComponent()

		' Add any initialization after the InitializeComponent() call.

	End Sub

	Public Event Minimized(ByVal Sender As frmChild)
	Public Event Restored(ByVal Sender As frmChild)

#Region " Form behaviour "

#Region " Resize, minimize, restore "

	Private Sub frmChild_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
		Static PreviousWindowState As System.Windows.Forms.FormWindowState = FormWindowState.Maximized
		If Not Me.Visible Then Return

		If Me.WindowState = FormWindowState.Minimized Then
			OnMinimized()
		Else
			If PreviousWindowState = FormWindowState.Minimized Then OnRestored()
		End If

		PreviousWindowState = Me.WindowState
	End Sub

	Private Sub frmChild_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		If Me.WindowState = FormWindowState.Minimized Then OnRestored()
	End Sub

	Protected Sub OnMinimized()
		RaiseEvent Minimized(Me)
	End Sub

	Protected Sub OnRestored()
		RaiseEvent Restored(Me)
	End Sub

#End Region

#End Region

End Class