Imports GuiLabs.Canvas.DrawStyle

Friend Class LineAppearance

    Implements ILineAppearance

    Private mLineStyle As ILineStyleInfo
    Public Property LineStyle() As ILineStyleInfo Implements DynamicGeometry.ILineAppearance.LineStyle
        Get
            Return mLineStyle
        End Get
        Set(ByVal Value As ILineStyleInfo)
            mLineStyle = Value
        End Set
    End Property

End Class
