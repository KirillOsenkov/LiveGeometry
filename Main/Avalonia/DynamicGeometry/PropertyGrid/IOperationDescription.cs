using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DynamicGeometry
{
    public interface IMetadataDescription
    {
        string Name { get; }
        string DisplayName { get; }
        T GetAttribute<T>() where T : Attribute;
        IEnumerable<T> GetAttributes<T>() where T : Attribute;
        object Parent { get; }
    }

    public interface IOperationDescription : IMetadataDescription
    {
        IEnumerable<IValueProvider> Parameters { get; }
        void Invoke(object target, IEnumerable<object> arguments);
    }

    public class MethodDescription : IOperationDescription, ICustomPropertyProvider
    {
        private readonly MethodInfo methodInfo;

        public MethodDescription(MethodInfo methodInfo)
        {
            this.methodInfo = methodInfo;
            TreatStaticMethodsAsInstance = true;
        }

        public bool TreatStaticMethodsAsInstance { get; set; }

        public static MethodDescription Get<T1>(string methodName)
        {
            var methodInfo = typeof(T1).GetMethod(methodName);
            if (methodInfo == null)
            {
                return null;
            }

            return new MethodDescription(methodInfo);
        }

        public static MethodDescription Create(MethodInfo methodInfo)
        {
            return new MethodDescription(methodInfo);
        }

        public IEnumerable<IValueProvider> Parameters
        {
            get
            {
                var result = methodInfo.GetParameters().Select(p => ValueProvider.Create(p, null));
                if (TreatStaticMethodsAsInstance && result.Any())
                {
                    result = result.Skip(1);
                }
                return result;
            }
        }

        public void Invoke(object target, IEnumerable<object> arguments)
        {
            var argumentList = arguments.ToList();
            if (methodInfo.IsStatic)
            {
                if (TreatStaticMethodsAsInstance)
                {
                    argumentList.Insert(0, target);
                }
                target = null;
            }
            methodInfo.Invoke(target, argumentList.ToArray());
        }

        public string Name
        {
            get
            {
                return methodInfo.Name;
            }
        }

        public string DisplayName
        {
            get
            {
                string name = Name;
                if (methodInfo.HasAttribute<PropertyGridNameAttribute>())
                {
                    name = methodInfo.GetAttribute<PropertyGridNameAttribute>().Name;
                }
                return name;
            }
        }

        public T GetAttribute<T>()
            where T : Attribute
        {
            return methodInfo.GetAttribute<T>();
        }

        public IEnumerable<T> GetAttributes<T>()
            where T : Attribute
        {
            return methodInfo.GetCustomAttributes(typeof(T), true).Cast<T>();
        }

        public object Parent
        {
            get
            {
                return null;
            }
        }

        public IEnumerable<IValueProvider> GetProperties()
        {
            return Parameters;
        }
    }
}
