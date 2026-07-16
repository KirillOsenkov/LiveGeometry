using System;

namespace DynamicGeometry
{
    public class DomainAttribute : Attribute
    {
        public DomainAttribute(double minValue, double maxValue)
        {
            MinValue = minValue;
            MaxValue = maxValue;
        }

        public double MinValue { get; set; }
        public double MaxValue { get; set; }
    }
}