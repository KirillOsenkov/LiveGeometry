using System;

namespace DynamicGeometry
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method,
        AllowMultiple = true, 
        Inherited = true)]
    public class PropertyGridEventAttribute : Attribute
    {
        public PropertyGridEventAttribute(string eventName, string handlerName)
        {
            EventName = eventName;
            HandlerName = handlerName;
        }

        public string EventName { get; private set; }
        public string HandlerName { get; private set; }
    }
}
