using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Markup;
using System.Windows.Media;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public interface ISerializer
    {
        object Write(IValueProvider value);
        void Read(IValueProvider value, string serialized);
        void Read(IValueProvider value, XElement serialized);
        bool CanSerialize(IValueProvider value);
        double Priority { get; }
    }

    public interface ISerializationService : ISerializer
    {
    }

    public static class SerializationServiceExtensions
    {
        public static T Read<T>(this ISerializationService service, XElement element)
        {
            ValueProvider value = new ValueProvider() { Type = typeof(T) };
            service.Read(value, element);
            return (T)value.Value;
        }
    }

    public class SerializationService : ISerializationService
    {
        public static SerializationService Instance { get; } = new SerializationService();

        public IEnumerable<ISerializer> Serializers { get; set; } = CollectSerializers();

        private static IEnumerable<ISerializer> CollectSerializers()
        {
            var assembly = typeof(SerializationService).Assembly;
            var types = assembly.GetTypes();
            var implementations = types.Where(t => 
                t.IsClass && 
                !t.IsAbstract &&
                t.HasInterface<ISerializer>() &&
                t != typeof(SerializationService));

            var instances = new List<ISerializer>();
            foreach (var type in implementations)
            {
                var instance = Activator.CreateInstance(type);
                if (instance is ISerializer serializer)
                {
                    instances.Add(serializer);
                }
            }

            return instances;
        }

        private Dictionary<Type, ISerializer> serializerCache = new Dictionary<Type, ISerializer>();

        private ISerializer FindBestSerializer(IValueProvider value)
        {
            ISerializer result = null;
            if (serializerCache.TryGetValue(value.Type, out result))
            {
                return result;
            }

            foreach (var serializer in Serializers.OrderBy(s => s.Priority))
            {
                if (serializer.CanSerialize(value))
                {
                    result = serializer;
                    serializerCache.Add(value.Type, result);
                    break;
                }
            }

            return result;
        }

        public object Write(IValueProvider value)
        {
            ISerializer serializer = FindBestSerializer(value);
            return serializer.Write(value);
        }

        public void Read(IValueProvider value, string serialized)
        {
            ISerializer serializer = FindBestSerializer(value);
            serializer.Read(value, serialized);
        }

        public void Read(IValueProvider value, XElement serialized)
        {
            ISerializer serializer = FindBestSerializer(value);
            serializer.Read(value, serialized);
        }

        public bool CanSerialize(IValueProvider value)
        {
            ISerializer serializer = FindBestSerializer(value);
            return serializer != null;
        }

        public double Priority
        {
            get
            {
                return double.MaxValue;
            }
        }
    }

    public class ComplexTypeSerializer : ISerializer
    {
        public object Write(IValueProvider value)
        {
            object instance = value.GetValue<object>();
            XElement result = new XElement(value.Name);

            var discoveryStrategy = ValueDiscoveryStrategy.Get(instance.GetType()) ?? valueDiscovery;
            var values = discoveryStrategy.GetValues(instance);
            foreach (var property in values)
            {
                object serialized = SerializationService.Instance.Write(property);

                if (serialized is string)
                {
                    XAttribute attribute = new XAttribute(property.Name, serialized.ToString());
                    result.Add(attribute);
                }
                else if (serialized is XElement)
                {
                    result.Add(new XElement(property.Name, serialized));
                }
            }

            return result;
        }

        public void Read(IValueProvider value, string serialized)
        {
            throw new NotImplementedException();
        }

        private IValueDiscoveryStrategy valueDiscovery = new IncludeByDefaultValueDiscoveryStrategy();

        public void Read(IValueProvider value, XElement serialized)
        {
            string name = serialized.Name.LocalName;
            var type = FindDerivedType(value.Type, name);
            var instance = Activator.CreateInstance(type);

            var discoveryStrategy = ValueDiscoveryStrategy.Get(type) ?? valueDiscovery;
            var values = discoveryStrategy.GetValues(instance);
            foreach (var property in values)
            {
                XAttribute attribute = serialized.Attribute(property.Name);
                if (attribute != null)
                {
                    SerializationService.Instance.Read(property, attribute.Value);
                    continue;
                }

                XElement subElement = serialized.Element(property.Name);
                if (subElement != null)
                {
                    SerializationService.Instance.Read(property, subElement);
                    continue;
                }
            }

            value.SetValue(instance);
        }

        private Type FindDerivedType(Type type, string name)
        {
            var assembly = typeof(ComplexTypeSerializer).Assembly;
            foreach (var candidate in assembly.GetTypes())
            {
                if (name == candidate.Name && type.IsAssignableFrom(candidate))
                {
                    return candidate;
                }
            }

            return null;
        }

        public bool CanSerialize(IValueProvider value)
        {
            return true;
        }

        public double Priority
        {
            get
            {
                return 3.0;
            }
        }
    }

    public abstract class SerializerBase<T> : ISerializer
    {
        public object Write(IValueProvider value)
        {
            return ToString(value.GetValue<T>());
        }

        public void Read(IValueProvider value, string serialized)
        {
            var deserialized = FromString(serialized);
            value.SetValue(deserialized);
        }

        public void Read(IValueProvider value, XElement serialized)
        {
            var deserialized = FromXElement(serialized);
            value.SetValue(deserialized);
        }

        public abstract string ToString(T value);

        public virtual T FromString(string str)
        {
            throw new NotImplementedException();
        }

        public virtual T FromXElement(XElement element)
        {
            throw new NotImplementedException();
        }

        public bool CanSerialize(IValueProvider value)
        {
            return value.Type == typeof(T);
        }

        public virtual double Priority
        {
            get
            {
                return 1.0;
            }
        }
    }

    public class StringSerializer : SerializerBase<string>
    {
        public override string FromString(string str)
        {
            return str;
        }

        public override string ToString(string value)
        {
            return value;
        }
    }

    public class FontFamilySerializer : SerializerBase<FontFamily>
    {
        public override FontFamily FromString(string str)
        {
            return new FontFamily(str);
        }

        public override string ToString(FontFamily value)
        {
            return value.Source;
        }
    }

    public class BoolSerializer : SerializerBase<bool>
    {
        public override bool FromString(string str)
        {
            return bool.Parse(str.ToLower());
        }

        public override string ToString(bool value)
        {
            return value ? "true" : "false";
        }
    }

    public class ColorSerializer : SerializerBase<Color>
    {
        public override string ToString(Color value)
        {
            return value.ToString();
        }

        public override Color FromString(string str)
        {
            return str.ToColor();
        }
    }

    public class BrushSerializer : ISerializer
    {
        public object Write(IValueProvider value)
        {
            object brush = value.GetValue<object>();
            SolidColorBrush solidColorBrush = brush as SolidColorBrush;
            if (solidColorBrush != null)
            {
                return solidColorBrush.Color.ToString();
            }

            return null;
        }

        public void Read(IValueProvider value, string serialized)
        {
            value.SetValue(new SolidColorBrush(serialized.ToColor()));
        }

        public void Read(IValueProvider value, XElement serialized)
        {
            serialized = serialized.Elements().First();
            SetNamespace(serialized, "http://schemas.microsoft.com/winfx/2006/xaml/presentation");
#if SILVERLIGHT
            var brush = XamlReader.Load(serialized.ToString());
#else
            var brush = XamlReader.Parse(serialized.ToString());
#endif
            value.SetValue(brush);
        }

        private void SetNamespace(XElement element, string ns)
        {
            element.Name = XName.Get(element.Name.LocalName, ns);
            foreach (var subElement in element.Elements())
            {
                SetNamespace(subElement, ns);
            }
        }

        public bool CanSerialize(IValueProvider value)
        {
            return value.Type == typeof(Brush);
        }

        public double Priority
        {
            get
            {
                return 1.0;
            }
        }
    }

    public class SimpleTypeSerializer : ISerializer
    {
        public object Write(IValueProvider value)
        {
            string result = "";
            var toStringMethod = GetToStringMethod(value);
            if (toStringMethod != null)
            {
                result = (string)toStringMethod.Invoke(value.GetValue<object>(), new object[] { CultureInfo.InvariantCulture });
            }
            else
            {
                result = value.GetValue<object>().ToString();
            }
            return result;
        }

        public void Read(IValueProvider value, string serialized)
        {
            var parseMethod = GetParseMethod(value);
            object result = null;
            if (parseMethod != null)
            {
                result = parseMethod.Invoke(null, new object[] { serialized, CultureInfo.InvariantCulture });
            }
            value.SetValue<object>(result);
        }

        public void Read(IValueProvider value, XElement serialized)
        {
            Read(value, serialized.Value);
        }

        static MethodInfo GetToStringMethod(IValueProvider value)
        {
            return value.Type.GetMethod("ToString", new[] { typeof(IFormatProvider) });
        }

        static MethodInfo GetParseMethod(IValueProvider value)
        {
            return value.Type.GetMethod("Parse", new[] { typeof(string), typeof(IFormatProvider) });
        }

        public bool CanSerialize(IValueProvider value)
        {
            return GetParseMethod(value) != null && GetToStringMethod(value) != null;
        }

        public double Priority
        {
            get
            {
                return 2.0;
            }
        }
    }
}