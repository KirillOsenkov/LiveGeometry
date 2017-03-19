using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;

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
                allStyles = style.StyleManager.GetSupportedStyles(figure);
            }

            Items = allStyles.Select(s => GetGlyph(s)).ToList();
            base.FillList();
        }

        protected override System.Windows.Controls.Primitives.Selector CreateSelector()
        {
            return new ListBox()
            {
                MaxHeight = 300
            };
        }

        private FrameworkElement GetGlyph(IFigureStyle s)
        {
            var content = s.GetSampleGlyph();
            content.Margin = new Thickness(10);
            var result = new Grid();
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
            if (value == null || Items == null)
            {
                return;
            }
            foreach (var item in Items)
            {
                if ((item as FrameworkElement).Tag == value)
                {
                    guard = true;
                    Selector.SelectedItem = item;
                    guard = false;
                }
            }
        }
    }
}
