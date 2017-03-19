Friend Interface IAppearanceFactory

    Function CreatePointAppearance(ByVal Color As System.Drawing.Color, ByVal FillColor As System.Drawing.Color, ByVal Shape As PointShapeType, ByVal Width As Integer) As IPointAppearance
    Function CreateLineAppearance(ByVal Color As System.Drawing.Color, ByVal DrawWidth As Integer) As ILineAppearance
    Function FindPointAppearance(ByVal StyleName As String) As IPointAppearance
    Function FindLineAppearance(ByVal StyleName As String) As ILineAppearance

End Interface
