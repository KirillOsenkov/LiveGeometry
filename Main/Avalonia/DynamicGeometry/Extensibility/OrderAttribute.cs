using System;

namespace DynamicGeometry
{
    public class OrderAttribute : Attribute
    {
        public double Order { get; set; }

        public OrderAttribute(double order)
        {
            Order = order;
        }
    }
}
