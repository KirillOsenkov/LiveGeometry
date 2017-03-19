using System.Collections.Generic;
using System.Windows;

namespace DynamicGeometry
{
    public class CompositeCommandObserver : ICommandObserver
    {
        private List<ICommandObserver> observers = new List<ICommandObserver>();

        public void CommandRemoved()
        {
            foreach (var observer in observers)
            {
                observer.CommandRemoved();
            }
        }

        public void EnabledChanged(bool newEnabledState)
        {
            for (int i = 0; i < observers.Count; i++)
            {
                observers[i].EnabledChanged(newEnabledState);
            }
        }

        public void Add(ICommandObserver observer)
        {
            observers.Add(observer);
        }

        public void IconChanged(FrameworkElement icon)
        {
            foreach (var observer in observers)
            {
                observer.IconChanged(icon);
            }
        }
    }
}
