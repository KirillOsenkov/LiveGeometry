using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DynamicGeometry
{
    public interface ICustomPropertyProvider
    {
        IEnumerable<IValueProvider> GetProperties();
    }

    public interface ICustomMethodProvider
    {
        IEnumerable<IOperationDescription> GetMethods();
    }

    public class CompositePropertyProvider : ICustomPropertyProvider
    {
        private IValueDiscoveryStrategy valueDiscoveryStrategy;
        private List<object> objects;

        public CompositePropertyProvider(IValueDiscoveryStrategy valueDiscoveryStrategy, IEnumerable<object> objects)
        {
            this.valueDiscoveryStrategy = valueDiscoveryStrategy;
            this.objects = new List<object>(objects ?? new object[0]);
        }

        Dictionary<string, CompositeValueProvider> results = new Dictionary<string, CompositeValueProvider>();

        public IEnumerable<IValueProvider> GetProperties()
        {
            if (objects.IsEmpty())
            {
                return null;
            }

            results.Clear();
            AddResults();

            return results
                .Where(k => k.Value.Count == objects.Count)
                .Select(k => (IValueProvider)k.Value);
        }

        void AddResults()
        {
            foreach (var item in objects)
            {
                var values = valueDiscoveryStrategy.GetValues(item);
                foreach (var value in values)
                {
                    AddResult(value);
                }
            }
        }

        void AddResult(IValueProvider value)
        {
            if (value.GetAttribute<PropertyGridDisallowMultiEditAttribute>() != null)
            {
                return;
            }
            var signature = GetSignature(value);

            CompositeValueProvider bucket;
            if (results.TryGetValue(signature, out bucket))
            {
                bucket.Add(value);
            }
            else
            {
                results.Add(signature, new CompositeValueProvider(value));
            }
        }

        string GetSignature(IValueProvider value)
        {
            return value.GetSignature();
        }

        public override string ToString()
        {
            if (objects.IsEmpty())
            {
                return "No objects selected";
            }
            else
            {
                return string.Format("{0} objects selected", objects.Count());
            }
        }
    }

    public interface IValueDiscoveryStrategy
    {
        IEnumerable<IValueProvider> GetValues(object editableObject);
    }

    public abstract class ValueDiscoveryStrategy : IValueDiscoveryStrategy
    {
        public abstract IEnumerable<IValueProvider> GetValues(object editableObject);

        public static IValueDiscoveryStrategy Get(Type type)
        {
            PropertyDiscoveryStrategyAttribute attribute = type.GetAttribute<PropertyDiscoveryStrategyAttribute>();
            if (attribute != null)
            {
                return (IValueDiscoveryStrategy)Activator.CreateInstance(attribute.ValueDiscoveryStrategyType);
            }

            return null;
        }
    }

    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface | AttributeTargets.Struct, Inherited = true, AllowMultiple = false)]
    public class PropertyDiscoveryStrategyAttribute : Attribute
    {
        public PropertyDiscoveryStrategyAttribute(Type type)
        {
            if (!typeof(IValueDiscoveryStrategy).IsAssignableFrom(type))
            {
                throw new ArgumentException
                    ("PropertyDiscoveryStrategyAttribute only accepts types that implement IValueDiscoveryStrategy");
            }
            ValueDiscoveryStrategyType = type;
        }

        public Type ValueDiscoveryStrategyType { get; set; }
    }

    public class PropertyDiscoveryStrategy : ValueDiscoveryStrategy
    {
        protected BindingFlags? BindingFlags { get; set; }

        protected virtual IEnumerable<PropertyInfo> GetProperties(object editableObject)
        {
            var type = editableObject.GetType();
            PropertyInfo[] properties;
            if (BindingFlags != null)
            {
                properties = type.GetProperties(BindingFlags.Value);
            }
            else
            {
                properties = type.GetProperties();
            }
            return properties
                .Where(p => p.GetIndexParameters().IsEmpty())
                //.OrderBy(p => p.Name)
                ;
        }

        protected virtual IEnumerable<PropertyInfo> FilterProperties(IEnumerable<PropertyInfo> properties)
        {
            return properties;
        }

        protected virtual IEnumerable<IValueProvider> CreateValueProviders(IEnumerable<PropertyInfo> properties, object editableObject)
        {
            foreach (var p in properties)
            {
                IValueProvider result = PropertyDiscoveryStrategy.CreateValueProvider(editableObject, p);
                yield return result;
            }
        }

        public override IEnumerable<IValueProvider> GetValues(object editableObject)
        {
            ICustomPropertyProvider provider = editableObject as ICustomPropertyProvider;
            if (provider != null)
            {
                return provider.GetProperties();
            }
            var properties = GetProperties(editableObject);
            properties = FilterProperties(properties);
            var result = CreateValueProviders(properties, editableObject);
            //if (editableObject is IEnumerable)
            //{
            //    int take = 10;
            //    var items = new List<IValueProvider>();
            //    foreach (var item in (IEnumerable)editableObject)
            //    {
            //        if (take-- <= 0)
            //        {
            //            break;
            //        }
            //        items.Add(new ObjectValue(item));
            //    }
            //    result = result.Concat(items);
            //}
            return result;
        }

        public static IValueProvider CreateValueProvider(object editableObject, PropertyInfo propertyInfo)
        {
            IValueProvider result = null;

            if (propertyInfo.HasAttribute<PropertyGridCustomValueProvider>())
            {
                var attribute = propertyInfo.GetAttribute<PropertyGridCustomValueProvider>();
                var derivedType = attribute.CustomValueProviderType;
                var instance = (PropertyValue)Activator.CreateInstance(derivedType);
                instance.Property = propertyInfo;
                instance.Parent = editableObject;
                instance.Type = propertyInfo.PropertyType;
                return instance;
            }
            else if (propertyInfo.PropertyType.HasInterface<IValueProvider>()
                && (result = (IValueProvider)propertyInfo.GetValue(editableObject, null)) != null)
            {
                return result;
            }
            else
            {
                return new PropertyValue(propertyInfo, editableObject);
            }
        }

        public static IValueProvider CreateValueProvider(object editableObject, string propertyName)
        {
            var propertyInfo = editableObject.GetType().GetProperty(propertyName);
            return CreateValueProvider(editableObject, propertyInfo);
        }

        public static IEnumerable<IValueProvider> GetValuesFromProperties(object editableObject, params string[] propertyNames)
        {
            foreach (var propertyName in propertyNames)
            {
                yield return CreateValueProvider(editableObject, propertyName);
            }
        }
    }

    public class IncludeByDefaultValueDiscoveryStrategy : PropertyDiscoveryStrategy
    {
        public static IncludeByDefaultValueDiscoveryStrategy Instance = new IncludeByDefaultValueDiscoveryStrategy();

        protected override IEnumerable<PropertyInfo> FilterProperties(IEnumerable<PropertyInfo> properties)
        {
            return properties
                .Where(p => !p.HasAttribute<IgnoreAttribute>() &&
                    (!p.HasAttribute<PropertyGridVisibleAttribute>()
                    || p.GetAttribute<PropertyGridVisibleAttribute>().Visible));
        }
    }

    public class ExcludeByDefaultValueDiscoveryStrategy : PropertyDiscoveryStrategy
    {
        public static ExcludeByDefaultValueDiscoveryStrategy Instance = new ExcludeByDefaultValueDiscoveryStrategy();

        protected override IEnumerable<PropertyInfo> FilterProperties(IEnumerable<PropertyInfo> properties)
        {
            return properties
                .Where(p =>
                    !p.HasAttribute<IgnoreAttribute>() &&
                    (p.HasAttribute<PropertyGridVisibleAttribute>()
                     && p.GetAttribute<PropertyGridVisibleAttribute>().Visible));
        }
    }
}