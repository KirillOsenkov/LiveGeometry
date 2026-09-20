using Avalonia.Controls;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    public class MethodCallerButton : Button
    {
        // Avalonia matches theme templates by concrete type; keep using the Button template.
        protected override System.Type StyleKeyOverride => typeof(Button);

        public MethodCallerButton()
        {
            this.Click += OnClick;
            this.Margin = new Avalonia.Thickness(0, 4, 0, 4);
        }

        private void OnClick(object sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (Target != null && OperationDescription != null)
            {
                try
                {
                    object[] arguments = new object[0];
                    if (ParameterGrid != null)
                    {
                        IEnumerable<IValueProvider> parameterValues = ParameterGrid.CurrentProperties;
                        arguments = parameterValues.Select(v => v.GetValue<object>()).ToArray();
                    }
                    OperationDescription.Invoke(Target, arguments);
                }
                catch
                {
                    // whatever happens, we can't allow it bubble up
                    // back to the CLR - we don't trust our method
                }
            }
        }

        /// <summary>
        /// A trash can to the left of the caption, both in the "careful" color.
        /// </summary>
        static object CreateDestructiveContent(string name)
        {
            var trashCan = new Avalonia.Controls.Shapes.Path()
            {
                // lid with handle, tapered body, two ribs - on a 14x14 grid
                Data = Avalonia.Media.Geometry.Parse(
                    "M2.5,4 H11.5 M5.5,4 V2.5 H8.5 V4 M3.5,4 L4.2,12 H9.8 L10.5,4 M6,6.2 V9.8 M8,6.2 V9.8"),
                Stroke = RibbonTheme.Destructive,
                StrokeThickness = 1.2,
                StrokeLineCap = Avalonia.Media.PenLineCap.Round,
                StrokeJoin = Avalonia.Media.PenLineJoin.Round,
                Width = 14,
                Height = 14,
                Margin = new Avalonia.Thickness(0, 0, 6, 0),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
            };
            var caption = new TextBlock()
            {
                Text = name,
                Foreground = RibbonTheme.Destructive,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
            };
            var result = new StackPanel() { Orientation = Avalonia.Layout.Orientation.Horizontal };
            result.Children.Add(trashCan);
            result.Children.Add(caption);
            return result;
        }

        public object Target { get; set; }
        public PropertyGrid ParameterGrid { get; set; }

        private IOperationDescription operationDescription;
        public IOperationDescription OperationDescription
        {
            get
            {
                return operationDescription;
            }
            set
            {
                operationDescription = value;
                if (operationDescription != null)
                {
                    string name = operationDescription.DisplayName;

                    var parameters = operationDescription.Parameters;
                    if (parameters.IsEmpty())
                    {
                        this.Content = operationDescription.GetAttribute<PropertyGridDestructiveAttribute>() != null
                            ? CreateDestructiveContent(name)
                            : name;
                    }
                    else
                    {
                        ParameterGrid = new PropertyGrid();
                        ParameterGrid.Title = name;
                        this.Content = ParameterGrid;
                        //this.HorizontalContentAlignment = Avalonia.HorizontalAlignment.Stretch;
                        //this.VerticalContentAlignment = Avalonia.VerticalAlignment.Stretch;
                        ParameterGrid.Show(operationDescription, null);
                    }
                }
            }
        }
    }
}
