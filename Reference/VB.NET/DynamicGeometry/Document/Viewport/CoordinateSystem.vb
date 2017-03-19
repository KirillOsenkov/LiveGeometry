Imports GuiLabs.Canvas
Imports GuiLabs.Canvas.Utils

Friend Class CoordinateSystem

    Implements ICoordinateSystem

	Public Sub New(ByVal NewParentPlane As IPlane)
		mParentPlane = NewParentPlane
		LoadIdentity()
		CacheTransformValues()
	End Sub

	Private WithEvents mParentPlane As IPlane
	Public ReadOnly Property ParentPlane() As DynamicGeometry.IPlane Implements DynamicGeometry.ICoordinateSystem.ParentPlane
		Get
			Return mParentPlane
		End Get
	End Property

#Region " LoadDefaultTransform "

    Private XOffset As Double = 0
    Private YOffset As Double = 0

    Private XScrOffset As Double = 0
    Private YScrOffset As Double = 0

    Private XScalar As Double = 1
    Private YScalar As Double = -1

    Private XUnit As Double = 0
    Private YUnit As Double = 0

    Private Const Epsilon As Double = 0.01

    Public Sub LoadDefaultTransform()
        LoadIdentity()
        CacheTransformValues()
        RefreshCanvasBorders()
    End Sub

    Public Sub LoadIdentity()
        XOffset = PixelsToUnits(ParentPlane.Width \ 2)
        YOffset = PixelsToUnits(ParentPlane.Height \ 2)

        XScrOffset = XOffset
        YScrOffset = YOffset

        XScalar = 1
        YScalar = -1

        XUnit = UnitsToPixels(1)
        YUnit = UnitsToPixels(1)
    End Sub

#Region " Cache transform values "

    Private XCache As Double = 0
    Private YCache As Double = 0

    Private XInvCache As Double = 0
    Private YInvCache As Double = 0

    Private XCachedOffset As Double = 0
    Private YCachedOffset As Double = 0

    Private XDiv As Double = 0
    Private YDiv As Double = 0

    Private Sub CacheTransformValues()
        XCache = XUnit * XScalar
        YCache = YUnit * YScalar

        XInvCache = 1 / XCache
        YInvCache = 1 / YCache

        XCachedOffset = XUnit * XOffset
        YCachedOffset = YUnit * YOffset

        XDiv = XOffset / XScalar
        YDiv = YOffset / YScalar
    End Sub

#End Region

#End Region

    '====================================================================

	Public Sub UpdateLogicalFromPhysical(ByVal toRecalc As CartesianPoint) Implements ICoordinateTransformer.UpdateLogicalFromPhysical
		toRecalc.SetLogical( _
		 toRecalc.X * XInvCache - XDiv, _
		 toRecalc.Y * YInvCache - YDiv)
	End Sub

	Public Sub UpdatePhysicalFromLogical(ByVal toRecalc As CartesianPoint) Implements ICoordinateTransformer.UpdatePhysicalFromLogical
		toRecalc.Set( _
		 CInt(toRecalc.UnitsX * XCache + XCachedOffset), _
		 CInt(toRecalc.UnitsY * YCache + YCachedOffset))
	End Sub

#Region " Translate units "

    Private Const PixelsPerUnit As Double = 32
    Public Function UnitsToPixels(ByVal Value As Double) As Long
        Return CLng(Value * PixelsPerUnit)
    End Function

    Public Function PixelsToUnits(ByVal Value As Long) As Double
        Return Value / PixelsPerUnit
    End Function

	'====================================================================

    Public Sub ToPhysical(ByRef x As Double, ByRef y As Double)
        x = x * XCache + XCachedOffset
        y = y * YCache + YCachedOffset
    End Sub

    '====================================================================

    Public Sub ToLogical(ByVal PhysicalPoint As Point, ByRef LogicalPoint As IMathPoint)
    End Sub

    Public Sub ToLogical(ByRef x As Double, ByRef y As Double)
        x = x * XInvCache - XDiv
        y = y * YInvCache - YDiv
    End Sub

#End Region

    '====================================================================

	Private CanvasBorders As MathTwoPoints
	Public ReadOnly Property Viewport() As MathTwoPoints Implements DynamicGeometry.ICoordinateSystem.Viewport
		Get
			Return CanvasBorders
		End Get
	End Property

    Private Sub mParentPlane_Resized(ByVal NewX As Integer, ByVal NewY As Integer) Handles mParentPlane.Resized
		If XOffset = 0 Or YOffset = 0 Then
			LoadIdentity()
			CacheTransformValues()
		End If

		RefreshCanvasBorders()
    End Sub

    Sub RefreshCanvasBorders()
        Dim x1 As Double = 0
        Dim y1 As Double = 0
        ToLogical(x1, y1)

        Dim x2 As Double = ParentPlane.Width
        Dim y2 As Double = ParentPlane.Height
        ToLogical(x2, y2)

		Common.SwapIfGreater(Of Double)(x1, x2)
		Common.SwapIfGreater(Of Double)(y1, y2)

		CanvasBorders.p1.x = x1
		CanvasBorders.p1.y = y1

		CanvasBorders.p2.x = x2
		CanvasBorders.p2.y = y2
    End Sub

End Class
