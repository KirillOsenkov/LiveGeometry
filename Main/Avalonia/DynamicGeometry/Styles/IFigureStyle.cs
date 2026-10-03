using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;

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

        /// <summary>The style as it looks under the theme on screen (see <see cref="FigureStyle.Resolve(string)"/>)</summary>
        IFigureStyle Resolve();
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

        // Asked for every figure that comes into a drawing (its default style), of every
        // style of the drawing until one fits: the attributes, read each time, were a good
        // part of loading a drawing
        static readonly ConcurrentDictionary<(Type StyleType, Type FigureType), bool> supports
            = new ConcurrentDictionary<(Type StyleType, Type FigureType), bool>();

        public static bool SupportsFigureType(this Type styleType, Type figureType)
        {
            return supports.GetOrAdd((styleType, figureType), types =>
            {
                foreach (var attribute in types.StyleType.GetAttributes<StyleForAttribute>())
                {
                    if (attribute.FigureBaseType.IsAssignableFrom(types.FigureType))
                    {
                        return true;
                    }
                }

                return false;
            });
        }

        public static void Apply(this FrameworkElement element, Style style)
        {
            if (style == null)
            {
                return;
            }
            style.ApplyTo(element);
        }

        public static void Apply(this IFigure figure, FrameworkElement element, IFigureStyle figureStyle)
        {
            var resolved = figureStyle.Resolve();
            var wpfStyle = resolved.GetWpfStyle(figure);
            element.Apply(wpfStyle);
            resolved.OnApplied(figure, element);
        }
    }
}