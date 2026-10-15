using System;

namespace LiveGeometry;

public class OrderAttribute : Attribute
{
    public double Order { get; set; }

    public OrderAttribute(double order)
    {
        Order = order;
    }
}
