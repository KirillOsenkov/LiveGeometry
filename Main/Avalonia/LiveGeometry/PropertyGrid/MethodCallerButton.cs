using Avalonia.Controls;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

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
                object[] arguments = [];
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
        return CreateContent(PropertyGridIcons.TrashCan(), name, nameof(AppTheme.Destructive));
    }

    /// <summary>
    /// An icon to the left of the caption, in a theme color.
    /// </summary>
    static object CreateContent(Control icon, string name, string foreground)
    {
        icon.Margin = new Avalonia.Thickness(0, 0, 6, 0);
        var caption = new TextBlock()
        {
            Text = name,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
        };
        caption.BindTheme(TextBlock.ForegroundProperty, foreground);
        var result = new StackPanel() { Orientation = Avalonia.Layout.Orientation.Horizontal };
        result.Children.Add(icon);
        result.Children.Add(caption);
        return result;
    }

    object CreateContent(IOperationDescription operation)
    {
        string name = operation.DisplayName;
        if (operation.GetAttribute<PropertyGridDestructiveAttribute>() != null)
        {
            return CreateDestructiveContent(name);
        }

        var iconAttribute = operation.GetAttribute<PropertyGridIconAttribute>();
        if (iconAttribute != null)
        {
            return CreateContent(PropertyGridIcons.Create(iconAttribute.Icon), name, nameof(AppTheme.Text));
        }

        return name;
    }

    public object Target { get; set; }
    public PropertyGrid ParameterGrid { get; set; }

    private IOperationDescription operationDescription;
    public IOperationDescription OperationDescription
    {
        get => operationDescription;
        set
        {
            operationDescription = value;
            if (operationDescription != null)
            {
                string name = operationDescription.DisplayName;

                var parameters = operationDescription.Parameters;
                if (parameters.IsEmpty())
                {
                    this.Content = CreateContent(operationDescription);
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
