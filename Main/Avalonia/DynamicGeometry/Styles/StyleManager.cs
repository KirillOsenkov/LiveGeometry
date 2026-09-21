using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Media;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public partial class StyleManager
    {
        static IEnumerable<Type> StyleTypes = Reflector.DiscoverTypes<IFigureStyle>();

        protected readonly ListWithEvents<IFigureStyle> list = new ListWithEvents<IFigureStyle>();

        public Drawing Drawing { get; set; }

        public StyleManager(Drawing drawing)
        {
            list.ItemAdded += OnStyleAdded;
            AddDefaultStyles();
            Drawing = drawing;
        }

        private void OnStyleAdded(IFigureStyle style)
        {
            if (style.Name.IsEmpty())
            {
                style.Name = GenerateUniqueName();
            }
            style.StyleManager = this;
        }

        protected virtual string GenerateUniqueName()
        {
            int n = 1;
            while (this[n.ToString()] != null)
            {
                n++;
            }
            return n.ToString();
        }

        public bool NameIsValid(string name)
        {
            return !list.Any(f => f.Name == name);
        }

        public IFigureStyle this[string index]
        {
            get
            {
                foreach (var style in list)
                {
                    if (style.Name == index)
                    {
                        return style;
                    }
                }

                return null;
            }
        }

        public IEnumerable<IFigureStyle> GetAllStyles()
        {
            return list;
        }

        public virtual IEnumerable<TStyle> GetStyles<TStyle>()
            where TStyle : class, IFigureStyle
        {
            foreach (var style in list)
            {
                TStyle correctStyle = style as TStyle;
                if (correctStyle != null)
                {
                    yield return correctStyle;
                }
            }
        }

        public IEnumerable<IFigureStyle> GetCompatibleStyles(Type styleType)
        {
            foreach (var style in list)
            {
                Type type = style.GetType();
                if (styleType == type)
                {
                    yield return style;
                }
            }
        }

        protected int numDefaultStyles;
        public virtual void AddDefaultStyles()
        {
            // The look tells how a point behaves: the two kinds that can be dragged are full
            // size and warm/bright, the constructed ones are a little smaller and cooler.
            var freePointStyle = new PointStyle()
            {
                Name = FreePointStyleName,
                Fill = new SolidColorBrush(Color.FromArgb(255, 255, 255, 100))
            };
            var pointOnFigureStyle = new PointStyle()
            {
                Name = PointOnFigureStyleName,
                Fill = new SolidColorBrush(Color.FromArgb(255, 124, 227, 139))
            };
            var intersectionPointStyle = new PointStyle()
            {
                Name = IntersectionPointStyleName,
                Size = 8,
                Fill = new SolidColorBrush(Color.FromArgb(255, 111, 211, 247))
            };
            var midpointStyle = new PointStyle()
            {
                Name = MidpointStyleName,
                Size = 8,
                Fill = new SolidColorBrush(Color.FromArgb(255, 255, 180, 90))
            };
            var dependentPointStyle = new PointStyle()
            {
                Name = DependentPointStyleName,
                Size = 8,
                Fill = new SolidColorBrush(Color.FromArgb(255, 208, 208, 208))
            };
            var lineStyle = new LineStyle();
            var lineStyle2 = new LineStyle()
            {
                Name = "OtherLine",
                Color = Color.FromArgb(200, 0, 0, 255)
            };
            var thickLineStyle = new LineStyle()
            {
                Name = "ThickLine",
                Color = Color.FromArgb(230, 0, 0, 0),
                StrokeWidth = 2.5
            };
            var redLineStyle = new LineStyle()
            {
                Name = "RedLine",
                Color = Color.FromArgb(255, 216, 59, 59),
                StrokeWidth = 1.5
            };
            var greenLineStyle = new LineStyle()
            {
                Name = "GreenLine",
                Color = Color.FromArgb(255, 46, 158, 79),
                StrokeWidth = 1.5
            };

            // auxiliary constructions: there, but stepping back
            var dashedLineStyle = new LineStyle()
            {
                Name = "DashedLine",
                Color = Color.FromArgb(255, 110, 110, 110),
                StrokeWidth = 1.25,
                Dash = LineDash.Dash
            };
            var dottedLineStyle = new LineStyle()
            {
                Name = "DottedLine",
                Color = Color.FromArgb(255, 110, 110, 110),
                StrokeWidth = 1.5,
                Dash = LineDash.Dot
            };

            // an outline with a hint of the same color inside: made for circles, fine for polygons
            var blueOutlineStyle = new ShapeStyle()
            {
                Name = "BlueOutline",
                Color = Color.FromArgb(255, 47, 123, 214),
                StrokeWidth = 1.5,
                Fill = new SolidColorBrush(Color.FromArgb(28, 47, 123, 214))
            };
            var orangeOutlineStyle = new ShapeStyle()
            {
                Name = "OrangeOutline",
                Color = Color.FromArgb(255, 224, 138, 0),
                StrokeWidth = 1.5,
                Fill = new SolidColorBrush(Color.FromArgb(28, 224, 138, 0))
            };
            var purpleOutlineStyle = new ShapeStyle()
            {
                Name = "PurpleOutline",
                Color = Color.FromArgb(255, 136, 84, 208),
                StrokeWidth = 1.5,
                Fill = new SolidColorBrush(Color.FromArgb(28, 136, 84, 208))
            };
            var shapeWithLineStyle = new ShapeStyle();
            var shapeStyle = new ShapeStyle()
            {
                Color = Colors.Transparent
            };
            var shapeStyle2 = new ShapeStyle()
            {
                Name = "OtherShape",
                Color = Colors.Transparent,
                Fill = new SolidColorBrush(Color.FromArgb(100, 200, 255, 200))
            };
            var hyperLinkStyle = new TextStyle()
            {
                FontSize = 18,
                FontFamily = new FontFamily("Segoe UI")
            };
            var textStyle = new TextStyle()
            {
                FontSize = 18,
                FontFamily = new FontFamily("Segoe UI")
            };
            var headerStyle = new TextStyle()
            {
                FontSize = 40,
                FontFamily = new FontFamily("Segoe UI")
            };

            var newStyles = new IFigureStyle[]
            {
                freePointStyle,
                pointOnFigureStyle,
                intersectionPointStyle,
                midpointStyle,
                dependentPointStyle,
                lineStyle,
                lineStyle2,
                thickLineStyle,
                redLineStyle,
                greenLineStyle,
                dashedLineStyle,
                dottedLineStyle,
                shapeStyle,
                shapeStyle2,
                shapeWithLineStyle,
                blueOutlineStyle,
                orangeOutlineStyle,
                purpleOutlineStyle,
                textStyle,
                headerStyle,
                hyperLinkStyle,
            };

            list.AddRange(newStyles);

            numDefaultStyles = newStyles.Length;
        }

        public IFigureStyle GetStyle(string name)
        {
            return GetAllStyles().Where(s => s.Name == name).FirstOrDefault();
        }

        public void SetStyleIfAvailable(IFigure figure, string styleName)
        {
            var style = GetStyle(styleName);
            if (style != null)
            {
                figure.Style = style;
            }
        }

        public IEnumerable<IFigureStyle> GetSupportedStyles(IFigure figure)
        {
            return GetSupportedStyles(figure.GetType());
        }

        public IEnumerable<IFigureStyle> GetSupportedStyles(Type figureType)
        {
            foreach (var style in list)
            {
                if (style.GetType().SupportsFigureType(figureType))
                {
                    yield return style;
                }
            }
        }

        public static Type GetStyleType(Type figureType)
        {
            foreach (var styleType in StyleTypes)
            {
                if (styleType.SupportsFigureType(figureType))
                {
                    return styleType;
                }
            }
            return null;
        }

        public const string FreePointStyleName = "FreePoint";
        public const string PointOnFigureStyleName = "PointOnFigure";
        public const string IntersectionPointStyleName = "IntersectionPoint";
        public const string MidpointStyleName = "Midpoint";

        // older drawings and PolygonIntersection know it under this name
        public const string DependentPointStyleName = "DependentPointStyle";

        public virtual IFigureStyle AssignDefaultStyle(IFigure figure)
        {
            // A drawing from a file brings its own styles; if it doesn't have the one for this
            // kind of point (older files don't), the first point style does, as before.
            var byKind = figure is IPoint ? GetStyle(GetDefaultPointStyleName(figure)) : null;
            if (byKind != null && byKind.GetType().SupportsFigureType(figure.GetType()))
            {
                return byKind;
            }

            var supportedStyles = GetSupportedStyles(figure);
            return supportedStyles.FirstOrDefault();
        }

        static string GetDefaultPointStyleName(IFigure point)
        {
            // PointOnFigure is a FreePoint, so it goes first
            if (point is PointOnFigure)
            {
                return PointOnFigureStyleName;
            }

            if (point is FreePoint)
            {
                return FreePointStyleName;
            }

            if (point is IntersectionPoint)
            {
                return IntersectionPointStyleName;
            }

            if (point is MidPoint)
            {
                return MidpointStyleName;
            }

            return DependentPointStyleName;
        }

        public IFigureStyle CreateNewStyle(IFigure figure)
        {
            var newStyle = figure.Style.Clone();
            Actions.AddItem(figure.Drawing.ActionManager, list, newStyle);
            return newStyle;
        }

        public IFigureStyle CreateNewStyle(Drawing drawing, IFigureStyle style)
        {
            var newStyle = style.Clone();
            Actions.AddItem(drawing.ActionManager, list, newStyle);
            return newStyle;
        }

        public IFigureStyle FindExistingOrAddNew(IFigureStyle style)
        {
            var found = list
                .Where(s => s.GetSignature() == style.GetSignature())
                .FirstOrDefault();
            if (found != null)
            {
                return found;
            }

            Actions.AddItem(Drawing.ActionManager, list, style);
            return style;
        }

        public void Clear()
        {
            list.Clear();
        }

        public virtual void Add(IFigureStyle style)
        {
            list.Add(style);
        }

        public virtual void Remove(IFigureStyle style)
        {
            if (list.Contains(style) && list.IndexOf(style) >= numDefaultStyles)
            {
                var transaction = Transaction.Create(Drawing.ActionManager, false);
                foreach (var fig in Drawing.Figures)
                {
                    if (fig.Style == style)
                    {
                        Actions.SetProperty(Drawing.ActionManager, new PropertyValue("Style",fig), AssignDefaultStyle(fig));
                    }
                }
                Actions.RemoveItem(Drawing.ActionManager, list, style);
                transaction.Commit();
            }
        }
    }
}