using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public class ComplexTypeEditorFactory : BaseValueEditorFactory<ComplexTypeEditor>
    {
        public ComplexTypeEditorFactory()
        {
            LoadOrder = 3;
        }

        public override bool SupportsValue(IValueProvider property)
        {
            return true;
        }
    }

    public class ComplexTypeEditor : PropertyGrid, IValueEditor
    {
        public ComplexTypeEditor()
        {
            ValueDiscoveryStrategy = new ExcludeByDefaultValueDiscoveryStrategy();
            Expanded = false;
            this.Margin = new Thickness(32, 0, 0, 0);
        }

        IValueProvider mValue;
        public IValueProvider Value
        {
            get
            {
                return mValue;
            }
            set
            {
                if (mValue == value)
                {
                    return;
                }
                if (mValue != null)
                {
                    mValue.ValueChanged -= mValue_ValueChanged;
                }
                mValue = value;
                if (mValue != null)
                {
                    mValue.ValueChanged += mValue_ValueChanged;
                    var attribute = value.GetAttribute<PropertyGridComplexTypeStateAttribute>();
                    if (attribute != null && attribute.State == ComplexTypeState.Expanded)
                    {
                        mExpanded = true;
                    }
                }
                OnValueSet(mValue);
            }
        }

        void mValue_ValueChanged()
        {
            OnValueSet(mValue);
        }

        void OnValueSet(IValueProvider value)
        {
            object selection = null;
            if (value != null)
            {
                selection = value.GetValue<object>();
            }
            Selection = selection;
        }

        const int iconSize = 10;
        const double glyphSize = 0.2;

        Canvas GetPlusIcon()
        {
            Canvas canvas = GetMinusIcon();
            Line line = new Line();
            line.X1 = 5;
            line.X2 = 5;
            line.Y1 = 3;
            line.Y2 = 8;
            canvas.Width = 10;
            canvas.Height = canvas.Width;
            canvas.Children.Add(line);
            return canvas;
        }

        Canvas GetMinusIcon()
        {
            Canvas canvas = new Canvas();
            Line line = new Line();
            line.X1 = 3;
            line.X2 = 8;
            line.Y1 = 5;
            line.Y2 = 5;
            canvas.Width = 10;
            canvas.Height = canvas.Width;
            canvas.Children.Add(line);
            return canvas;
        }

        Border expandCollapse;

        protected override void AddHeader()
        {
            StackPanel header = new StackPanel();
            header.Margin = new Thickness(-32, 0, 0, 0);
            header.Orientation = Orientation.Horizontal;
            header.VerticalAlignment = VerticalAlignment.Center;
            expandCollapse = new Border();
            expandCollapse.Background = new SolidColorBrush(Colors.White);
            expandCollapse.BorderBrush = new SolidColorBrush(Colors.Black);
            expandCollapse.BorderThickness = new Thickness(1);
            expandCollapse.MouseLeftButtonDown += expandCollapse_Click;
            UpdateExpandCollapseGlyph();
            expandCollapse.HorizontalAlignment = HorizontalAlignment.Center;
            expandCollapse.VerticalAlignment = VerticalAlignment.Center;
            expandCollapse.Margin = new Thickness(4);
            expandCollapse.Padding = new Thickness();
            header.Children.Add(expandCollapse);

            TextBlock name = new TextBlock();
            name.Margin = new Thickness(4, 0, 4, 0);
            name.VerticalAlignment = VerticalAlignment.Center;
            name.Text = Value.DisplayName;
            header.Children.Add(name);

            object value = Value.GetValue<object>();
            string contentsText = (value ?? "").ToString();
            if (!string.IsNullOrEmpty(contentsText) && value != null && contentsText != value.GetType().ToString())
            {
                TextBlock contents = new TextBlock();
                contents.MaxWidth = 200;
                contents.MaxHeight = 30;
                contents.Margin = new Thickness(16, 0, 0, 0);
                contents.VerticalAlignment = VerticalAlignment.Center;
                contents.Text = contentsText;
                header.Children.Add(contents);
            }
            this.Children.Add(header);
        }

        protected override void AddChildren()
        {
            var controls = CreateObjectControls(Selection);
            if (controls.IsEmpty())
            {
                expandCollapse.Visibility = Visibility.Collapsed;
                return;
            }
            foreach (var control in controls)
            {
                this.Children.Add(control);
            }
        }

        void expandCollapse_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Expanded = !Expanded;
            UpdateExpandCollapseGlyph();
        }

        void UpdateExpandCollapseGlyph()
        {
            expandCollapse.Child = Expanded ? GetMinusIcon() : GetPlusIcon();
        }
    }
}