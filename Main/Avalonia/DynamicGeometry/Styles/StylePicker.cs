using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Layout;

namespace DynamicGeometry
{
    public class StylePickerEditorFactory : BaseValueEditorFactory<StylePickerEditor>
    {
        public override bool SupportsValue(IValueProvider value)
        {
            return value.Type == typeof(IFigureStyle) && base.SupportsValue(value);
        }
    }

    public class StylePropertyValueProvider : PropertyValue
    {
        public override string GetSignature()
        {
            string result = Name + CanSetValue.ToString() + Type.ToString();
            if (Name == "StyleDisplay" && Parent != null)
            {
                result += StyleManager.GetStyleType(Parent.GetType());
            }
            return result;
        }
    }

    public class StylePickerEditor : SelectorValueEditor, IValueEditor
    {
        /// <summary>
        /// 1. Get the type of the style being displayed
        /// 2. Get the list of all other styles like this one
        /// 3. Fill all the styles in the list
        /// </summary>
        public override void FillList()
        {
            IFigureStyle style = Value.GetValue<IFigureStyle>();
            if (style == null && Value is CompositeValueProvider)
            {
                style = (Value as CompositeValueProvider).InnerList[0].GetValue<IFigureStyle>();
            }

            if (style == null)
            {
                base.FillList();
                return;
            }

            IEnumerable<IFigureStyle> allStyles = style.GetCompatibleStyles();

            IFigure figure = Value.Parent as IFigure;
            if (figure != null && style != null)
            {
                allStyles = style.StyleManager.GetSupportedStyles(figure)
                    .Where(candidate => StyleManager.IsOffered(candidate, figure));
            }

            Items = allStyles.Select(s => GetGlyph(s)).ToList();
            base.FillList();
        }

        /// <summary>
        /// The side of the square every sample sits in, whatever its size: the picker is a
        /// grid, eight to a row (the list is at most 340 wide, and an item takes 3 more), as
        /// many as the default styles of a kind have in a row (<see cref="StyleHue"/>)
        /// </summary>
        public const double CellSize = 36;

        protected override Selector CreateSelector()
        {
            return new ListBox()
            {
                MaxHeight = 300
            };
        }

        private FrameworkElement GetGlyph(IFigureStyle s)
        {
            var content = s.GetSampleGlyph();
            content.HorizontalAlignment = HorizontalAlignment.Center;
            content.VerticalAlignment = VerticalAlignment.Center;
            var result = new Grid()
            {
                Width = CellSize,
                Height = CellSize
            };
            result.Children.Add(content);
            result.Tag = s;
            return result;
        }

        protected override ValidationResult Validate(object value)
        {
            IFigureStyle style = (value as FrameworkElement).Tag as IFigureStyle;
            var result = new ValidationResult();
            if (style != null)
            {
                result.IsValid = true;
                result.Value = style;
            }
            return result;
        }

        public override void UpdateEditor()
        {
            var value = GetValue();
            ShowSelected(item => value != null && (item as FrameworkElement).Tag == value);
        }
    }
}
