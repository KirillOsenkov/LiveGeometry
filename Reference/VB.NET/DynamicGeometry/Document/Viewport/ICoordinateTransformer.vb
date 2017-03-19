Imports GuiLabs.Canvas

Friend Interface ICoordinateTransformer

	Sub UpdateLogicalFromPhysical(ByVal toRecalc As CartesianPoint)
	Sub UpdatePhysicalFromLogical(ByVal toRecalc As CartesianPoint)

End Interface