using System.ComponentModel;

namespace DynamicGeometry
{
    public interface INotifyPropertyChanging
    {
        event PropertyChangedEventHandler PropertyChanging;
    }
}