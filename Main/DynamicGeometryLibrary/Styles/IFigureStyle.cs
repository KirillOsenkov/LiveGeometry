using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace DynamicGeometry
{
    [PropertyDiscoveryStrategy(typeof(ExcludeByDefaultValueDiscoveryStrategy))]
    public interface IFigureStyle : INotifyPropertyChanged
    {
        string Name { get; set; }
        StyleManager StyleManager { get; set; }
        Style GetWpfStyle(IFigure figure);
        FrameworkElement GetSampleGlyph();
        
        /// <summary>
        /// A style's signature is a string that can be used
        /// to compare two styles. If two signatures are equal,
        /// then two styles are equal as well.
        /// </summary>
        string GetSignature();
        IFigureStyle Clone();
        IEnumerable<IFigureStyle> GetCompatibleStyles();
        void OnApplied(IFigure figure, FrameworkElement element);
#if !PLAYER
        FigureStyle.EditInfo CurrentEditInfo { get; set; }
#endif
    }

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false)]
    public class StyleForAttribute : Attribute
    {
        public StyleForAttribute(Type figureType)
        {
            FigureBaseType = figureType;
        }

        public Type FigureBaseType { get; set; }
    }

    public static class StyleExtensions
    {
        public static Style GetWpfStyle(this IFigureStyle style)
        {
            return style.GetWpfStyle(null);
        }

        public static bool SupportsFigureType(this Type styleType, Type figureType)
        {
            var attributes = styleType.GetAttributes<StyleForAttribute>();
            foreach (var attribute in attributes)
            {
                if (attribute.FigureBaseType.IsAssignableFrom(figureType))
                {
                    return true;
                }
            }
            return false;
        }

        public static void Apply(this FrameworkElement element, Style style)
        {
            if (style == null)
            {
                return;
            }
            foreach (Setter setter in style.Setters)
            {
                element.ClearValue(setter.Property);
            }
            element.Style = style;
        }

        public static void Apply(this IFigure figure, FrameworkElement element, IFigureStyle figureStyle)
        {
            var wpfStyle = figureStyle.GetWpfStyle(figure);
            element.Apply(wpfStyle);
            figureStyle.OnApplied(figure, element);
        }
    }
}