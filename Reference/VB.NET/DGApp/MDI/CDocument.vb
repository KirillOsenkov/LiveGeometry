Imports DynamicGeometry

Public Class CDocument

	Public Event Closed(ByVal Doc As CDocument)

	'====================================================================
	' Create a new document and a new MDI child window to host it.
	'====================================================================
	Public Sub New(ByVal ExistingParentForm As frmParent)
		ParentForm = ExistingParentForm
		View = DynamicGeometry.Facade.CreateDGView()
        Drawing = DynamicGeometry.Facade.CreateDGDocument
        View.Document = Drawing
        HostChildForm = CreateNewMDIChild()

		InitializeForm()
		InitializeMenu()
		Drawing.CmdFactory.GetCommand(DynamicGeometry.CommandStrings.Mode.Drag).Checked = True

		HostChildForm.Show()
		AddHandler System.Windows.Forms.Application.Idle, AddressOf OnIdle
	End Sub

	Public Sub OnIdle(ByVal sender As Object, ByVal e As System.EventArgs)
		Drawing.CmdFactory.UpdateCommands()
	End Sub

	Public Function CreateNewMDIChild() As frmChild
		Dim NewForm As New frmChild()
		NewForm.MdiParent = ParentForm
		AddHandler NewForm.Minimized, AddressOf ParentForm.OnChildMinimized
		AddHandler NewForm.Restored, AddressOf ParentForm.OnChildRestored
		AddHandler NewForm.Enter, AddressOf OnChildEnter
		Return NewForm
	End Function

	Private Sub InitializeForm()
		HostChildForm.SuspendLayout()

		HostChildForm.Controls.Add(View)

		View.Dock = System.Windows.Forms.DockStyle.Fill
		View.Name = "Canvas"
		View.BackColor = Color.AliceBlue

		HostChildForm.ResumeLayout(False)
	End Sub

	Private Sub InitializeMenu()
		Dim DGControlFactory As ControlFactory = New ControlFactory(Me.Drawing.CmdFactory)
		Dim MenuBuilder As IMenuBuilder = New DGMenuBuilder()
		Dim DGTBuilder As IToolbarBuilder = New DGToolbarBuilder()

		HostChildForm.Menu = MenuBuilder.BuildMenu(DGControlFactory)
		HostChildForm.tlbGeometry.Buttons.AddRange(DGTBuilder.BuildToolbar(DGControlFactory))
	End Sub

	'====================================================================

	Private mParentForm As frmParent
	Public Property ParentForm() As frmParent
		Get
			Return mParentForm
		End Get
		Set(ByVal Value As frmParent)
			mParentForm = Value
		End Set
	End Property

	Private mDrawing As IDGDocument
	Public Property Drawing() As IDGDocument
		Get
			Return mDrawing
		End Get
		Set(ByVal Value As IDGDocument)
			mDrawing = Value
		End Set
	End Property

	Private WithEvents mHostChildForm As frmChild
	Public Overridable Property HostChildForm() As frmChild
		Get
			Return mHostChildForm
		End Get
		Set(ByVal Value As frmChild)
			mHostChildForm = Value
		End Set
	End Property

    Private mView As DGView
    Public Property View() As DGView
        Get
            Return mView
        End Get
        Set(ByVal Value As DGView)
            mView = Value
        End Set
    End Property

	Private Sub mHostChildForm_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles mHostChildForm.Closed
		RaiseEvent Closed(Me)
	End Sub

	Public Sub OnChildEnter(ByVal Sender As Object, ByVal e As System.EventArgs)
		HostChildForm.Width = HostChildForm.Width
	End Sub

End Class