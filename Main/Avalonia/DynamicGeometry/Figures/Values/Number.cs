using System;

namespace DynamicGeometry
{
    public class Number : FigureBase, INumber
    {
        public override void ApplyStyle()
        {
            
        }

        public override IFigure HitTest(Avalonia.Point point)
        {
            return null;
        }

        public Func<double> Function { get; set; }
        public string Text { get; set; }

        public double Value
        {
            get
            {
                if (Function != null)
                {
                    return Function();
                }
                return 0;
            }
        }
    }
}
