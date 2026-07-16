using System;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;

namespace DynamicGeometry
{
    public static class Reflector
    {
        public static List<Type> DiscoverTypes<T>()
        {
            var result = new List<Type>();
            foreach (var type in Assembly.GetExecutingAssembly().GetTypes())
            {
                if (typeof(T).IsAssignableFrom(type) && CanInstantiate(type))
                {
                    result.Add(type);
                }
            }
            return result;
        }

        public static List<T> DiscoverTypesAndInstantiate<T>()
        {
            var result = new List<T>();
            foreach (var type in Assembly.GetExecutingAssembly().GetTypes())
            {
                if (typeof(T).IsAssignableFrom(type) && CanInstantiate(type))
                {
                    result.Add((T)Activator.CreateInstance(type));
                }
            }
            return result;
        }

        public static bool CanInstantiate(Type type)
        {
            return (type.IsClass || type.IsValueType)
                && !type.IsAbstract;
        }

        public static EventInfo FindEventByName(Type type, string eventName)
        {
            var events = type.GetEvents();
            return events.FirstOrDefault(e => e.Name == eventName);
        }

        public static MethodInfo FindMethodByName(Type type, string methodName)
        {
            var methods = type.GetMethods();
            return methods.FirstOrDefault(m => m.Name == methodName);
        }
    }
}