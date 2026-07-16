using System.Collections.Generic;

namespace DynamicGeometry
{
    public class LineTwoPoints : LineBase, ILine
    {
        public override PointPair OnScreenCoordinates
        {
            get
            {
                return Math.GetLineFromSegment(Coordinates, CanvasLogicalBorders);
            }
        }

        public static void Convert(ILine oldLine, ILine newLine)
        {
            var drawing = oldLine.Drawing;
            newLine.Style = oldLine.Style;
            Actions.ReplaceWithNew(oldLine, newLine);
            drawing.RaiseUserIsAddingFigures(new Drawing.UIAFEventArgs() { Figures = newLine.AsEnumerable<IFigure>() });
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to ray")]
        public void ConvertToRay()
        {
            LineTwoPoints.Convert(this, Factory.CreateRay(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to segment")]
        public void ConvertToSegment()
        {
            LineTwoPoints.Convert(this, Factory.CreateSegment(this.Drawing, this.Dependencies));
        }

#endif

    }
}