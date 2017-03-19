'Imports GuiLabs.Canvas.Renderer

'Friend Class PointShape

'    Implements IPointShape

'    Public Sub New(ByVal x As Integer, ByVal y As Integer)
'        MoveTo(x, y)
'    End Sub

'    Public Sub New()

'    End Sub

'    Public Sub Draw(ByVal Renderer As IRenderer) Implements IPointShape.Draw

'    End Sub


'    Public Sub MoveTo(ByVal NewX As Integer, ByVal NewY As Integer) Implements DynamicGeometry.IPointShape.MoveTo
'        x = NewX
'        y = NewY
'        'mBoundRect.Location = New System.Drawing.Point()
'        'PhysicalX = x
'        'PhysicalY = y
'        'Appearance.mRect.X = PhysicalX - Appearance.mRadius
'        'Appearance.mRect.Y = PhysicalY - Appearance.mRadius

'        'If ShouldRefresh Then
'        '	'ParentList.Refresh()
'        'End If
'    End Sub

'    'Public Sub Move(ByVal deltaX As Integer, ByVal deltaY As Integer, Optional ByVal ShouldRefresh As Boolean = False)
'    '	'SetPosition(PhysicalX + deltaX, PhysicalY + deltaY, ShouldRefresh)
'    'End Sub

'    Public Function HitTest(ByVal x As Integer, ByVal y As Integer) As Boolean Implements DynamicGeometry.IPointShape.HitTest


'        'Return PixelShape.BoundRect.Contains(mx, my)

'        'mx = Math.Abs(mx - PixelShape.x)
'        'my = Math.Abs(my - PixelShape.y)
'        'If mx > my Then my = mx
'        'Return my <= Appearance.mSensibleRadius
'        'Return Math.Abs(mx - PhysicalX) + Math.Abs(my - PhysicalY) <= Appearance.mSensibleRadius
'        'Return Math.Sqrt(mx * mx + my * my) <= Appearance.mSensibleRadius

'    End Function

'End Class