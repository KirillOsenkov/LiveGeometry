namespace DynamicGeometry
{
    public class SegmentCreator : FigureCreator
    {
        protected override ExpectedDependencyList InitExpectedDependencies()
        {
            return ExpectedDependencyList.PointPoint;
        }

        protected override IFigure CreateFigure()
        {
            LineTwoPoints result = Factory.CreateLineTwoPoints(FoundDependencies);
            return result;
        }

        public override string Icon
        {
            get
            {
                return "resources/bitmaps/geometry%20toolbar/dgwsegment.bmp";
            }
        }
    }
}