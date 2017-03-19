Friend Interface IFigure

	' Figures that contain this figure.
	' The list normally remains empty. 
	' When we add this segment to a triangle as a new side of it, 
	' we should remember this triangle, because we are now a part of it
	'Property ContainingFigures() As IFigureList


	' Figures, which are parts of current complex (component) figure
	' this figure contains some other figures: they are called Components (Parts)
	'Property Components() As IFigureList
	' TODO: TOTHINK: I think we don't need to explicitly publish a list of components for each type of figure
	' it can perfectly remain private for the one special figure type: 
	' CompositeFigure (aka Macro aka complex figure aka ensemble)
	' Recommendation: remove this property from IFigure interface


	' prerequisites, required to create current figure
	' that is, direct parents, from which current figure depends
	Property Parents() As IFigureList


	' direct children, that depend on current figure
	Property Children() As IFigureList

	ReadOnly Property FigureType() As IFigureType

	Sub Recalculate()

End Interface
