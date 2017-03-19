using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Shapes;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Media.Effects;

namespace DynamicGeometry
{
    public static class CanvasExtensions
    {
        public static void MoveTo(this FrameworkElement element, Point position)
        {
            MoveTo(element, position.X, position.Y);
        }

        public static void MoveTo(this FrameworkElement element, double x, double y)
        {
            Canvas.SetLeft(element, x - element.Width / 2);
            Canvas.SetTop(element, y - element.Height / 2);
        }
    }

    public class Factory
    {
        public static Shape CreatePointShape()
        {
            //FrameworkElement result = new PointShape();
            //container.Children.Add(result);
            //return result;

            int size = 8;
            Ellipse ellipse = new Ellipse()
            {
                Width = size,
                Height = size,
                Fill = new SolidColorBrush(Colors.LightYellow),
                BitmapEffect = new DropShadowBitmapEffect()
                {
                    ShadowDepth = 4,
                    Opacity = 0.7
                    //Softness = 0.5
                },
                Stroke = Brushes.Black,
                StrokeThickness = 0.5
            };

            return ellipse;
        }

        public static Line CreateLineShape()
        {
            Line result = new Line()
            {
                Stroke = Brushes.Black,
                StrokeThickness = 0.7
            };

            return result;
        }

        public static MidPoint CreateMidPoint(IFigureList dependencies)
        {
            MidPoint result = new MidPoint(dependencies);
            return result;
        }

        public static LineTwoPoints CreateLineTwoPoints(IFigureList dependencies)
        {
            return new LineTwoPoints(dependencies);
        }

        public static FreePoint CreateFreePoint()
        {
            FreePoint result = new FreePoint();
            return result;
        }
    }
}
