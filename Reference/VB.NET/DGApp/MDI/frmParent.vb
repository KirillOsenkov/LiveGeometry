Imports DynamicGeometry

Public Class frmParent

	Inherits System.Windows.Forms.Form

	' New & Dispose call GlobalThings.Initialize & .Dispose; global initialization region
#Region " Windows Form Designer generated code "

	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
		MyBase.Dispose(disposing)
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.

	Friend WithEvents imlMainbar As System.Windows.Forms.ImageList
	Friend WithEvents tlbMainbar As DGApp.DGToolbar

	Friend WithEvents sbStatus As System.Windows.Forms.StatusBar

	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmParent))
		Me.imlMainbar = New System.Windows.Forms.ImageList(Me.components)
		Me.tlbMainbar = New DGApp.DGToolbar()
		Me.sbStatus = New System.Windows.Forms.StatusBar()
		Me.SuspendLayout()
		'
		'imlMainbar
		'
		Me.imlMainbar.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
		Me.imlMainbar.ImageSize = New System.Drawing.Size(16, 16)
		Me.imlMainbar.ImageStream = CType(resources.GetObject("imlMainbar.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.imlMainbar.TransparentColor = System.Drawing.Color.Magenta
		'
		'tlbMainbar
		'
		Me.tlbMainbar.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
		Me.tlbMainbar.DropDownArrows = True
		Me.tlbMainbar.ImageList = Me.imlMainbar
		Me.tlbMainbar.Name = "tlbMainbar"
		Me.tlbMainbar.ShowToolTips = True
		Me.tlbMainbar.Size = New System.Drawing.Size(512, 39)
		Me.tlbMainbar.TabIndex = 6
		'
		'sbStatus
		'
		Me.sbStatus.Location = New System.Drawing.Point(0, 394)
		Me.sbStatus.Name = "sbStatus"
		Me.sbStatus.Size = New System.Drawing.Size(512, 22)
		Me.sbStatus.TabIndex = 9
		'
		'frmParent
		'
		Me.ClientSize = New System.Drawing.Size(512, 416)
		Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.sbStatus, Me.tlbMainbar})
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.IsMdiContainer = True
		Me.MinimumSize = New System.Drawing.Size(200, 200)
		Me.Name = "frmParent"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "DG"
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.ResumeLayout(False)

	End Sub

#End Region

	Public Sub New()
		MyBase.New()

		mCmdFactory = New DGAppCommandFactory(Me)

		'This call is required by the Windows Form Designer.
		InitializeComponent()
		InitializeMenu()
		AddHandler System.Windows.Forms.Application.Idle, AddressOf OnIdle
	End Sub

	Public Sub OnIdle(ByVal sender As Object, ByVal e As System.EventArgs)
		CmdFactory.UpdateCommands()
	End Sub

	Private Sub frmParent_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
		CmdFactory.GetCommand(CommandStrings.FileNew).OnClick()
		If Documents IsNot Nothing Then
			Documents(0).HostChildForm.WindowState = FormWindowState.Maximized
		End If
	End Sub

	'====================================================================

	Protected Sub InitializeMenu()
		Dim AppControlFactory As ControlFactory = New ControlFactory(CmdFactory)

		Dim MAMBuilder As IMenuBuilder = New AppMenuBuilder()
		Dim TBuilder As IToolbarBuilder = New AppToolbarBuilder()

		Me.Menu = MAMBuilder.BuildMenu(AppControlFactory)
		Me.tlbMainbar.Buttons.AddRange(TBuilder.BuildToolbar(AppControlFactory))
	End Sub

	Private mCmdFactory As ICommandFactory
	Public ReadOnly Property CmdFactory() As ICommandFactory
		Get
			Return mCmdFactory
		End Get
	End Property

	Public Function GetActiveDocument() As CDocument
		Dim CDoc As CDocument
		Dim ActiveForm As frmChild = DirectCast(Me.ActiveMdiChild, frmChild)
		For Each CDoc In Documents
			If CDoc.HostChildForm Is ActiveForm Then
				Return CDoc
			End If
		Next
		Return Nothing
	End Function

	'Private mDGCmdFactory As ICommandFactory
	'Public ReadOnly Property DGCmdFactory() As ICommandFactory
	'	Get
	'		Return mDGCmdFactory
	'	End Get
	'End Property

#Region " Documents() "

	Private WithEvents mDocuments As New CDocumentCollection()
	Public ReadOnly Property Documents() As CDocumentCollection
		Get
			Return mDocuments
		End Get
	End Property

	'Public ReadOnly Property ActiveDocument() As CDocument
	'	Get
	'		Dim w As frmChild = Me.ActiveMdiChild
	'		If w Is Nothing Then Return Nothing
	'		Return w.Doc
	'	End Get
	'End Property

	'Public Sub CloseDocument()
	'	Dim m As frmChild = Me.ActiveMdiChild		  ' Find active MDI child window
	'	If Not IsNothing(m) Then m.Close() ' if there is any, close it
	'End Sub

#End Region

#Region " Interface OUT "
	' ===============================================================================================================

	Public Sub EnableMainToolbar(ByVal Enable As Boolean)
		tlbMainbar.Enabled = Enable
	End Sub

	Public Sub EnableMenu(ByVal Enable As Boolean)
		CmdFactory.GetCommand(CommandStrings.File).Enabled = Enable
		CmdFactory.GetCommand(CommandStrings.Window).Enabled = Enable
	End Sub

	Public Sub Status(ByVal Text As String)
		sbStatus.Text = Text
	End Sub

	' ===============================================================================================================
#End Region

#Region " MDI menu functionality"	' ===============================================================================================================

	Protected Sub OnFirstDocumentCreated() Handles mDocuments.CollectionNotMoreEmpty
		SetMenuVisibility(True)
	End Sub

	Protected Sub OnLastDocumentDeleted() Handles mDocuments.CollectionEmptied
		SetMenuVisibility(False)
	End Sub

	Protected Sub SetMenuVisibility(ByVal B As Boolean)
		CmdFactory.GetCommand(CommandStrings.FileClose).Visible = B
		CmdFactory.GetCommand(CommandStrings.Window).Visible = B
	End Sub

	Public NumberOfMinimizedChildWindows As Integer = 0

	Public Sub OnChildMinimized(ByVal Sender As frmChild)
		If Me.NumberOfMinimizedChildWindows = 0 Then
			CmdFactory.GetCommand(CommandStrings.WindowArrange).Visible = True
		End If
		Me.NumberOfMinimizedChildWindows += 1
	End Sub

	Public Sub OnChildRestored(ByVal Sender As frmChild)
		Me.NumberOfMinimizedChildWindows -= 1
		If Me.NumberOfMinimizedChildWindows = 0 Then
			CmdFactory.GetCommand(CommandStrings.WindowArrange).Visible = False
		End If
	End Sub

	' ===============================================================================================================
#End Region	' any open MDIChildren? Change the UI Accordingly

	'#Region " Interface IN "
	'	' ===============================================================================================================
	'	'Private Sub frmParent_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
	'		'LockWindowUpdate(Me.Handle.ToInt32)
	'		'MenuCommand(MenuCommandID.cmdNew)
	'		'DirectCast((New CommandFileNew()), ICommand).OnClick()
	'		'Me.Visible = True
	'		'LockWindowUpdate(0)
	'	'End Sub

	'	'#Region " Toolbars "
	'	'	' ===============================================================================================================
	'	'	Private Sub tlbMainbar_ButtonClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles tlbMainbar.ButtonClick
	'	'		MenuCommand(MenuCommandID.cmdNew.Parse(GetType(MenuCommandID), e.Button.Tag, True))
	'	'	End Sub

	'	'	Public Sub tlbGeometry_ToolChangedByUser(ByVal NewTool As GeometryTool) Handles tlbGeometry.ToolChangedByUser
	'	'		Dim d As CDocument = App.ActiveDocument
	'	'		If Not d Is Nothing Then
	'	'			d.CurrentTool = NewTool
	'	'		End If
	'	'	End Sub

	'	'	Public Property ActiveTool() As GeometryTool
	'	'		Get
	'	'			Return tlbGeometry.ActiveTool
	'	'		End Get
	'	'		Set(ByVal Value As GeometryTool)
	'	'			tlbGeometry.ActiveTool = Value
	'	'		End Set
	'	'	End Property
	'	'	' ===============================================================================================================
	'	'#End Region

	'#End Region ' menu, toolbars, window events (frmParent_Load)

End Class