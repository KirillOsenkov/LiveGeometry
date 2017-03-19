Imports GuiLabs.Canvas.Renderer

Friend MustInherit Class DGLine

    Inherits DGObject
    Implements IDGLine

	Public Sub New(ByVal ContainerDoc As DGDocument, ByVal NewP1 As IDGPoint, ByVal NewP2 As IDGPoint)
		MyBase.New(ContainerDoc)
		Parents.Add(NewP1)
		Parents.Add(NewP2)
		Me.Style = AppearanceFactory.Instance.CreateLineAppearance(System.Drawing.Color.Black, 1)
	End Sub

	Private mStyle As ILineAppearance
    Public Property Style() As ILineAppearance
        Get
            Return mStyle
        End Get
        Set(ByVal Value As ILineAppearance)
            mStyle = Value
        End Set
    End Property

    Public Property P1() As IDGPoint Implements IDGLine.P1
        Get
            Return DirectCast(Parents(0), IDGPoint)
        End Get
        Set(ByVal Value As IDGPoint)
            Parents(0) = Value
        End Set
    End Property

    Public Property P2() As IDGPoint Implements IDGLine.P2
        Get
            Return DirectCast(Parents(1), IDGPoint)
        End Get
        Set(ByVal Value As IDGPoint)
            Parents(1) = Value
        End Set
    End Property

	Public Overridable Sub Draw(ByVal CurrentRenderer As IRenderer) Implements DynamicGeometry.IDGObject.Draw
		CurrentRenderer.DrawOperations.DrawLine(Me.P1.Coordinates, Me.P2.Coordinates, Me.Style.LineStyle)
	End Sub

    Public MustOverride Function IsPointOver(ByVal x As Integer, ByVal y As Integer) As Boolean Implements DynamicGeometry.IDGObject.IsPointOver

End Class
