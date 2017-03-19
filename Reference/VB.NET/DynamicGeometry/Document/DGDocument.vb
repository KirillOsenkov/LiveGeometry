Imports System.Drawing
Imports System.Windows.Forms
Imports GuiLabs.Canvas.Events
Imports GuiLabs.Canvas.Renderer
Imports DynamicGeometry.Actions

#Region " Enum GeometryTool, UIMode "

Public Enum GeometryTool
	Pointer
	BasePoint
	Segment
End Enum

Public Enum UIMode
	Normal
	CreateFigure
End Enum

#End Region

Friend Class DGDocument
    Inherits KeyMouseHandler
    Implements IDGDocument

    Public Sub New()
        mCmdFactory = New DGCommandFactory(Me)
        ActiveTool = CommandStrings.Mode.Drag
        Behaviour = New Dragger(Me)
    End Sub

#Region " Events "

    Public Event NeedRedraw() Implements IDGDocument.NeedRedraw

#End Region

    Public Sub Draw(ByVal Renderer As IRenderer) Implements IDGDocument.Draw
        Behaviour.Draw(Renderer)
    End Sub

    Public Sub RaiseNeedRedraw()
		RaiseEvent NeedRedraw()
	End Sub

    '=================================================================================================================================
    ' The Document's data:
    '=================================================================================================================================

    ' List of Figures
    Private mFigures As VisualObjectList = New VisualObjectList()
    Public Property Figures() As VisualObjectList
        Get
            Return mFigures
        End Get
        Set(ByVal Value As VisualObjectList)
            mFigures = Value
        End Set
    End Property

    ' Add, remove, modify, etc. actions, that can be done with current document
    Private mActionManager As ActionCenter = New ActionCenter(Me)
    Public ReadOnly Property ActionManager() As ActionCenter
        Get
            Return mActionManager
        End Get
	End Property

	Private mCmdFactory As CommandFactory
	Public ReadOnly Property CmdFactory() As ICommandFactory Implements IDGDocument.CmdFactory
		Get
			Return mCmdFactory
		End Get
	End Property

	Private mActiveTool As String
	Public Property ActiveTool() As String
		Get
			Return mActiveTool
		End Get
		Set(ByVal Value As String)
			mActiveTool = Value
		End Set
	End Property

	' Mathematical drawing plane
	Private mPaper As IPlane = New Plane()
	Public Property Paper() As IPlane
		Get
			Return mPaper
		End Get
		Set(ByVal Value As IPlane)
			mPaper = Value
		End Set
	End Property

	Public Sub Resize(ByVal x As Integer, ByVal y As Integer) Implements IDGDocument.Resize
		Paper.Resize(x, y)
	End Sub

	Private mBehaviour As IBehaviour
	Public Property Behaviour() As IBehaviour
		Get
			Return mBehaviour
		End Get
		Set(ByVal Value As IBehaviour)
			mBehaviour = Value
			Me.DefaultMouseHandler = Behaviour
		End Set
	End Property

	Public Overrides Sub OnMouseDown(ByVal e As GuiLabs.Canvas.Events.MouseEventArgsWithKeys)
		MyBase.OnMouseDown(e)
	End Sub

#Region " Handling View events " '====================================================================

	'Private Sub mView_Repaint() Handles mView.Repaint
	'    Behaviour.Repaint()
	'End Sub

	'Private Sub mView_DoubleClick(ByVal e As DynamicGeometry.MouseEventArgsWithKeys) Handles mView.DoubleClick
	'    Behaviour.DoubleClick(e)
	'End Sub

	'Private Sub mView_MouseDown(ByVal e As DynamicGeometry.MouseEventArgsWithKeys) Handles mView.MouseDown
	'    Behaviour.MouseDown(e)
	'End Sub

	'Private Sub mView_MouseMove(ByVal e As DynamicGeometry.MouseEventArgsWithKeys) Handles mView.MouseMove
	'    Behaviour.MouseMove(e)
	'End Sub

	'Private Sub mView_MouseUp(ByVal e As DynamicGeometry.MouseEventArgsWithKeys) Handles mView.MouseUp
	'    Behaviour.MouseUp(e)
	'End Sub

	'Private Sub mView_MouseWheel(ByVal e As DynamicGeometry.MouseEventArgsWithKeys) Handles mView.MouseWheel
	'    Behaviour.MouseWheel(e)
	'End Sub

	'Private Sub mView_MouseHover() Handles mView.MouseHover
	'    Behaviour.MouseHover()
	'End Sub

#End Region	' mouse events and paint - redirection from View to Behaviour '=============

	'Private WithEvents mView As IDGView
	'Public Property View() As IDGView
	'	Get
	'		Return mView
	'	End Get
	'	Set(ByVal Value As IDGView)
	'		mView = Value
	'		Figures.Canvas = mView.Canvas
	'	End Set
	'End Property

	'Public WriteOnly Property ViewControl() As System.Windows.Forms.Control Implements DynamicGeometry.IDGDocument.ViewControl
	'	Set(ByVal Value As System.Windows.Forms.Control)
	'		If TypeOf (Value) Is IDGView Then
	'			View = DirectCast(Value, IDGView)
	'		Else
	'			' TODO: Add exception-handling
	'		End If
	'	End Set
	'End Property

End Class