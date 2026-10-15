using System.ComponentModel;

namespace LiveGeometry;

public interface INotifyPropertyChanging
{
    event PropertyChangedEventHandler PropertyChanging;
}