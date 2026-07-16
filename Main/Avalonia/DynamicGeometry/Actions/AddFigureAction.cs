namespace DynamicGeometry
{
    public class AddFigureAction : GeometryAction
    {
        public AddFigureAction(Drawing drawing, IFigure figure)
            : base(drawing)
        {
            Figure = figure;
        }

        public IFigure Figure { get; set; }

        protected override void ExecuteCore()
        {
            Drawing.Figures.Add(Figure);
        }

        protected override void UnExecuteCore()
        {
            Drawing.Figures.Remove(Figure);
        }
    }
}
