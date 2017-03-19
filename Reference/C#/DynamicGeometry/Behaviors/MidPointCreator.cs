namespace DynamicGeometry
{
    public class MidpointCreator : FigureCreator
    {
        protected override ExpectedDependencyList InitExpectedDependencies()
        {
            return ExpectedDependencyList.PointPoint;
        }

        protected override IFigure CreateFigure()
        {
            MidPoint result = Factory.CreateMidPoint(FoundDependencies);
            return result;
        }

        public override string Icon
        {
            get
            {
                return "resources/bitmaps/geometry%20toolbar/dgwmidpoint.bmp";
            }
        }
    }
}