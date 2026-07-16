using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DynamicGeometry
{
    public interface IValueProvider : IMetadataDescription
    {
        event Action ValueChanged;
        void RaiseValueChanged();
        T GetValue<T>();
        bool CanSetValue { get; }
        void SetValue<T>(T value);
        Type Type { get; }
        string GetSignature();
    }

    public class CompositeValueProvider : IValueProvider
    {
        public List<IValueProvider> InnerList = new List<IValueProvider>();

        public CompositeValueProvider(IEnumerable<IValueProvider> providers)
        {
            foreach (var item in providers)
            {
                Add(item);
            }
        }

        public CompositeValueProvider(IValueProvider value)
        {
            Add(value);
        }

        public void Add(IValueProvider provider)
        {
            InnerList.Add(provider);
            provider.ValueChanged += RaiseValueChanged;
        }

        public int Count
        {
            get
            {
                return InnerList.Count;
            }
        }

        public event Action ValueChanged;

        public void RaiseValueChanged()
        {
            if (ValueChanged != null)
            {
                ValueChanged();
            }
        }

        public T GetValue<T>()
        {
            T commonValue = InnerList[0].GetValue<T>();
            if (commonValue == null)
            {
                return commonValue;
            }
            foreach (var item in InnerList)
            {
                if (!commonValue.Equals(item.GetValue<T>()))
                {
                    return default(T);
                }
            }
            return commonValue;
        }

        public bool CanSetValue
        {
            get { return InnerList.All(i => i.CanSetValue); }
        }

        public void SetValue<T>(T value)
        {
            foreach (var item in InnerList)
            {
                item.SetValue(value);
            }
        }

        public object Parent
        {
            get 
            {
                object commonParent = InnerList[0].Parent;
                foreach (var item in InnerList)
                {
                    if (commonParent != item.Parent)
                    {
                        return null;
                    }
                }
                return commonParent;
            }
        }

        public Type Type
        {
            get { return InnerList[0].Type; }
        }

        public string Name
        {
            get { return InnerList[0].Name; }
        }

        public string DisplayName
        {
            get { return InnerList[0].DisplayName; }
        }

        public T GetAttribute<T>() where T : Attribute
        {
            return InnerList[0].GetAttribute<T>();
        }

        public IEnumerable<T> GetAttributes<T>() where T : Attribute
        {
            return InnerList[0].GetAttributes<T>();
        }

        public string GetSignature()
        {
            return InnerList.ToListString(",");
        }
    }

    public class ObjectValue : IValueProvider
    {
        public ObjectValue(object instance)
        {
            Instance = instance;
            Type = instance.GetType();
        }

        public object Instance { get; set; }
        public Type Type { get; private set; }

        public T GetValue<T>()
        {
            return (T)Instance;
        }

        public bool CanSetValue
        {
            get { return true; }
        }

        public void SetValue<T>(T value)
        {
            Instance = value;
        }

        public string DisplayName
        {
            get { return Instance.ToString(); }
        }

        public T GetAttribute<T>() where T : Attribute
        {
            return default(T);
        }

        public IEnumerable<T> GetAttributes<T>() where T : Attribute
        {
            return Enumerable.Empty<T>();
        }

        public object Parent
        {
            get { return null; }
        }

        public string Name
        {
            get { return DisplayName; }
        }

        public event Action ValueChanged;

        public void RaiseValueChanged()
        {
            if (ValueChanged != null)
            {
                ValueChanged();
            }
        }

        public virtual string GetSignature()
        {
            string result = Name + CanSetValue.ToString() + Type.ToString();
            return result;
        }
    }

    public class ValueProvider : IValueProvider
    {
        public object Parent { get; set; }
        public Type Type { get; set; }
        public object Value { get; set; }

        public event Action ValueChanged;
        public void RaiseValueChanged()
        {
            if (ValueChanged != null)
            {
                ValueChanged();
            }
        }

        public virtual T GetValue<T>()
        {
            return (T)Value;
        }

        public virtual bool CanSetValue
        {
            get { return true; }
        }

        public virtual void SetValue<T>(T value)
        {
            Value = value;
        }

        public virtual string Name { get; set; }

        public virtual string DisplayName { get; set; }

        public virtual T GetAttribute<T>() where T : Attribute
        {
            return null;
        }

        public virtual IEnumerable<T> GetAttributes<T>() where T : Attribute
        {
            return Enumerable.Empty<T>();
        }

        public virtual string GetSignature()
        {
            string result = Name + CanSetValue.ToString() + Type.ToString();
            return result;
        }

        public static IValueProvider Create(ParameterInfo parameterInfo, object parent)
        {
            var result = new ValueProvider()
            {
                Type = parameterInfo.ParameterType,
                Name = parameterInfo.Name,
                Parent = parent,
                DisplayName = parameterInfo.Name,
            };

            return result;
        }
    }

    public class PropertyValue : ValueProvider
    {
        public PropertyValue()
        {
        }

        public PropertyValue(PropertyInfo propertyInfo, object instance)
        {
            Property = propertyInfo;
            Parent = instance;
            Type = propertyInfo.PropertyType;
        }

        public PropertyValue(string propertyName, object instance)
        {
            Parent = instance;
            Type = instance.GetType();
            Property = Type.GetProperty(propertyName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        }

        public PropertyInfo Property { get; set; }

        public override T GetValue<T>()
        {
            return (T)Property.GetValue(Parent, null);
        }

        public override string Name
        {
            get
            {
                return Property.Name;
            }
        }

        string mName = null;
        public override string DisplayName
        {
            get
            {
                if (mName == null)
                {
                    mName = Name;
                    if (Property.HasAttribute<PropertyGridNameAttribute>())
                    {
                        mName = Property.GetAttribute<PropertyGridNameAttribute>().Name;
                    }
                }
                return mName;
            }
        }

        public override string ToString()
        {
            return DisplayName;
        }

        public override bool CanSetValue
        {
            get { return Property.CanWrite; }
        }

        public override void SetValue<T>(T value)
        {
            Property.SetValue(Parent, value, null);
        }

        public override T GetAttribute<T>()
        {
            return Property.GetAttribute<T>();
        }

        public override IEnumerable<T> GetAttributes<T>()
        {
            return Property.GetCustomAttributes(typeof(T), true).Cast<T>();
        }
    }
}