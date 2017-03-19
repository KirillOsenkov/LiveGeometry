using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Globalization;

namespace DynamicGeometry
{
    public class IgnoreAttribute : Attribute { }

    public static class Utilities
    {
        public static string ToListString<T>(this IEnumerable<T> list, string separator)
        {
            if (list == null)
            {
                return null;
            }
            var strings = list.Select(i => i.ToString()).ToArray();
            var result = string.Join(separator, strings);
            return result;
        }

        public static bool HasInterface<T>(this Type type)
        {
            return typeof(T).IsAssignableFrom(type);
        }

        public static double Round(this double num, int fractionalDigits)
        {
            return Math.Round(num, fractionalDigits);
        }

        public static string ToStringInvariant(this double number)
        {
            return number.ToString(CultureInfo.InvariantCulture);
        }

        public static double Round(this double num)
        {
            return Math.Round(num);
        }

        public static bool IsEmpty<T>(this IEnumerable<T> collection)
        {
            return collection == null || !collection.Any();
        }

        public static bool IsEmpty<T>(this IList<T> collection)
        {
            return collection == null || collection.Count == 0;
        }

        public static bool IsEmpty(this string s)
        {
            return string.IsNullOrEmpty(s);
        }

        public static bool HasAttribute<T>(
            this MemberInfo attributeHost)
            where T : Attribute
        {
            return attributeHost.GetAttribute<T>() != null;
        }

        public static T GetAttribute<T>(
            this MemberInfo attributeHost)
            where T : Attribute
        {
            return (T)Attribute.GetCustomAttribute(attributeHost, typeof(T));
        }

        public static IEnumerable<T> GetAttributes<T>(
            this MemberInfo attributeHost)
            where T : Attribute
        {
            return Attribute.GetCustomAttributes(attributeHost, typeof(T)).OfType<T>();
        }
    }
}
