Imports DynamicGeometry

Public Class DGToolbar
	Inherits System.Windows.Forms.ToolBar

	Public Sub New()
		MyBase.New()
	End Sub

	Protected Overrides Sub OnButtonClick(ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs)
		'Try
		Dim Button As DGToolBarButton = DirectCast(e.Button, DGToolBarButton)
		Button.UICommand.OnClick()

		'Catch ex As Exception
		'	MsgBox(ex.Message)
		'End Try
	End Sub
End Class

'====================================================================
'#Region " old toolbar "

'Public Class DGToolbarOld
'	Inherits System.Windows.Forms.ToolBar

'	Public Sub New()
'		MyBase.New()
'	End Sub

'#Region " Tool map " ' ==================================================================================================================
'	' one-to-one correspondence between buttons and GeometryTool constants
'	Public GeometryToolMap As New IntegerMap()

'#Region " Bijection: buttons <=> GeometryTool constants " ' ==================================================================================================================
'	' strictly speaking, it is an injection from buttons to GeometryTool constants;
'	' you could potentially hide some buttons so that some constants get unused

'	Public Function Tool(ByVal Index As Integer) As GeometryTool
'		Return GeometryToolMap.XofY(Index)
'	End Function

'	Public Function ToolIndex(ByVal aTool As GeometryTool) As Integer
'		Return GeometryToolMap.YofX(aTool)
'	End Function
'	' ==================================================================================================================
'#End Region

'	' Physical setting of the "pushed" state
'	Private Sub SetPushedState(ByVal aTool As GeometryTool, Optional ByVal NewPushedState As Boolean = True)
'		Me.Buttons(ToolIndex(aTool)).Pushed = NewPushedState
'	End Sub

'	' ==================================================================================================================
'#End Region

'#Region " Interface OUT (ActiveTool property) " ' ==================================================================================================================
'	Public Event ToolChangedByUser(ByVal NewTool As GeometryTool)

'	Private mActiveTool As GeometryTool = GeometryTool.BasePoint
'	Public Property ActiveTool() As GeometryTool
'		Get
'			Return mActiveTool
'		End Get
'		Set(ByVal Value As GeometryTool)
'			If mActiveTool = Value Then Return

'			' This ensures ONE AND ONLY ONE tool can be active at a time
'			SetPushedState(mActiveTool, False)
'			mActiveTool = Value
'			SetPushedState(mActiveTool)
'		End Set
'	End Property
'	' ==================================================================================================================
'#End Region

'#Region " Interface IN (OnButtonClick) "
'	Protected Overrides Sub OnButtonClick(ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs)

'		Dim pressedButton As Integer = Me.Buttons.IndexOf(e.Button)
'		Dim newTool As GeometryTool = Me.Tool(pressedButton)

'		If newTool = ActiveTool Then
'			If newTool = GeometryTool.Pointer Then
'				SetPushedState(GeometryTool.Pointer)
'			Else
'				ActiveTool = GeometryTool.Pointer
'				RaiseEvent ToolChangedByUser(GeometryTool.Pointer)
'			End If
'		Else
'			ActiveTool = newTool
'			RaiseEvent ToolChangedByUser(newTool)
'		End If
'	End Sub
'#End Region
'End Class
'#End Region
