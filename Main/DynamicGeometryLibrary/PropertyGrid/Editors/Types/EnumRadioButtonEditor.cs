using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class EnumRadioButtonEditorFactory : BaseValueEditorFactory<EnumRadioButtonEditor>
    {
        public EnumRadioButtonEditorFactory()
        {
            LoadOrder = 2;
        }

        public override bool SupportsValue(IValueProvider value)
        {
            return value.Type.IsEnum && base.SupportsValue(value);
        }
    }

    public class EnumRadioButtonEditor : LabeledValueEditor, IValueEditor
    {
        private Border groupBox;
        private StackPanel stackPanel;
        private IEnumerable<string> items;

        protected override UIElement CreateEditor()
        {
            groupBox = new Border();
            groupBox.BorderThickness = new Thickness(1);
            groupBox.BorderBrush = new SolidColorBrush(Colors.LightGray);
            stackPanel = new StackPanel();
            stackPanel.Margin = new Thickness(7);
            groupBox.Child = stackPanel;

            return groupBox;
        }

        protected override void InitCore()
        {
            var propertyType = Value.Type;
            items = from f in propertyType.GetFields()
                    where f.FieldType == propertyType
                    select Enum.Parse(propertyType, f.Name, true).ToString();

            bool first = true;
            foreach (string enumField in items)
            {
                RadioButton radioButton = new RadioButton();
                radioButton.Content = enumField;
                radioButton.Margin = new Thickness(0, 5, 0, 0);
                if (first)
                {
                    first = false;
                    radioButton.Margin = new Thickness(0);
                }
                radioButton.Checked += RadioButton_Checked;
                radioButton.Tag = Enum.Parse(Value.Type, enumField, true);
                stackPanel.Children.Add(radioButton);
            }

            UpdateEditor();
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var value = (sender as RadioButton).Tag;
            SetValue(value);
        }

        public override void UpdateEditor()
        {
            object value = Value.GetValue<object>();
            foreach (RadioButton item in stackPanel.Children)
            {
                if (value.Equals(item.Tag))
                {
                    item.IsChecked = true;
                    return;
                }
            }
        }

        protected override ValidationResult Validate(object value)
        {
            return new ValidationResult()
            {
                IsValid = true,
                Value = Enum.Parse(Value.Type, value.ToString(), false)
            };
        }
    }
}
