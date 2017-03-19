Friend Class Plane

    Implements IPlane

    Public Event Resized(ByVal NewX As Integer, ByVal NewY As Integer) Implements IPlane.Resized

    Public Sub Resize(ByVal NewX As Integer, ByVal NewY As Integer) Implements DynamicGeometry.IPlane.Resize
        Width = NewX
        Height = NewY
        RaiseEvent Resized(NewX, NewY)
    End Sub

    Private mHeight As Integer
    Public Property Height() As Integer Implements DynamicGeometry.IPlane.Height
        Get
            Return mHeight
        End Get
        Set(ByVal Value As Integer)
            mHeight = Value
        End Set
    End Property

    Private mWidth As Integer
    Public Property Width() As Integer Implements DynamicGeometry.IPlane.Width
        Get
            Return mWidth
        End Get
        Set(ByVal Value As Integer)
            mWidth = Value
        End Set
    End Property

    Private mActiveCoordinateSystem As ICoordinateSystem = New CoordinateSystem(Me)
    Public ReadOnly Property ActiveCoordinateSystem() As DynamicGeometry.ICoordinateSystem Implements DynamicGeometry.IPlane.ActiveCoordinateSystem
        Get
            Return mActiveCoordinateSystem
        End Get
    End Property

End Class
