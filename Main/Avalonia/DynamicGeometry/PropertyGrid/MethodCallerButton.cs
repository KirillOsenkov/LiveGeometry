using System.Reflection;
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
                        this.Content = name;
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
