using System;
using System.Windows;

namespace DynamicGeometry
{
    public class Command : ICommand
    {
#if !PLAYER
        public DrawingHost DrawingHost { get; set; }
#endif
        public Command()
        {
            Observers = new CompositeCommandObserver();
        }

        public Command(Action implementation, FrameworkElement icon, string name, string category) : this()
        {
            Implementation = implementation;
            Icon = icon;
            Name = name;
            Category = category;
        }
#if !PLAYER
        public Command(DrawingHost drawingHost, Action implementation, FrameworkElement icon, string name, string category)
            : this(implementation,icon,name,category)
        {
            DrawingHost = drawingHost;
        }
#endif
        bool enabled = true;
        public bool Enabled
        {
            get
            {
                return enabled;
            }
            set
            {
                enabled = value;
                Observers.EnabledChanged(value);
            }
        }

        protected Action Implementation { get; set; }

        public CompositeCommandObserver Observers { get; set; }

        FrameworkElement icon;
        public FrameworkElement Icon
        {
            get
            {
                return icon;
            }
            set
            {
                icon = value;
                Observers.IconChanged(value);
            }
        }

        public string Name { get; set; }
        public string Category { get; set; }

        public virtual void Execute()
        {
            if (Implementation != null)
            {
                Implementation();
            }
#if !PLAYER
            if (DrawingHost != null)
            {
                DrawingHost.RaiseCommandExecuted(this);
            }
#endif
        }

        public void AddObserver(ICommandObserver observer)
        {
            Observers.Add(observer);
        }
    }
}
