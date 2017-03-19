using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
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
            var freePointStyle = new PointStyle();
            var pointOnFigureStyle = new PointStyle()
                {
                    Fill = new SolidColorBrush(Color.FromArgb(255, 0, 255, 0))
                };
            var dependentPointStyle = new PointStyle()
                {
                    Fill = new SolidColorBrush(Color.FromArgb(255, 192, 192, 192))
                };
            var lineStyle = new LineStyle();
            var shapeWithLineStyle = new ShapeStyle();
            var shapeStyle = new ShapeStyle()
            {
                Color = Colors.Transparent
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

            (new IFigureStyle[]
            { 
                freePointStyle,
                pointOnFigureStyle,
                dependentPointStyle,
                lineStyle,
                shapeStyle,
                shapeWithLineStyle,
                textStyle,
                headerStyle,
                hyperLinkStyle,
            }).ForEach(list.Add);
            numDefaultStyles = 9;
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

        public virtual IFigureStyle AssignDefaultStyle(IFigure figure)
        {
            var supportedStyles = GetSupportedStyles(figure);
            return supportedStyles.FirstOrDefault();
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