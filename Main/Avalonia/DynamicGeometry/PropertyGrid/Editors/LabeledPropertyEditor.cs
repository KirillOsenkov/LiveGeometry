using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public abstract partial class LabeledValueEditor : Grid
    {
        public LabeledValueEditor()
        {
            this.ColumnDefinitions.Add(new ColumnDefinition()
            {
                Width = GridLength.Auto,
                MinWidth = 60,
                // all rows of a property grid share the width of the label column
                SharedSizeGroup = "PropertyGridLabel"
            });
            this.ColumnDefinitions.Add(new ColumnDefinition());
            this.HorizontalAlignment = HorizontalAlignment.Stretch;
            Label = new TextBlock();
            Label.VerticalAlignment = VerticalAlignment.Center;
            Label.Margin = new Thickness(2, 0, 12, 0);
            Label.SetValue(Grid.ColumnProperty, 0);
            var editor = CreateEditor();
            editor.SetValue(Grid.ColumnProperty, 1);
            this.Children.Add(Label);
            this.Children.Add(editor);
            this.Loaded += LabeledPropertyEditor_Loaded;
        }

        void LabeledPropertyEditor_Loaded(object sender, RoutedEventArgs e)
        {
            if (Value != null && Value.GetAttribute<PropertyGridFocusAttribute>() != null)
            {
                Focus();
            }
        }

        protected virtual void Focus()
        {

        }

        protected abstract UIElement CreateEditor();

        public TextBlock Label { get; set; }

        IValueProvider mValue;
        public IValueProvider Value
        {
            get
            {
                return mValue;
            }
            set
            {
                if (mValue != null)
                {
                    mValue.ValueChanged -= mValue_ValueChanged;
                }
                mValue = value;
                if (mValue != null)
                {
                    mValue.ValueChanged += mValue_ValueChanged;
                    OnValueSet();
                }
            }
        }

        public ActionManager ActionManager { get; set; }

        protected virtual void mValue_ValueChanged()
        {
            if (guard)
            {
                return;
            }
            guard = true;
            UpdateEditor();
            guard = false;
        }

        private bool guard = false;
        protected virtual void OnValueSet()
        {
            if (guard)
            {
                return;
            }
            guard = true;
            Label.Text = Value.DisplayName;
            InitCore();
            UpdateEditor();
            guard = false;
        }

        protected virtual void InitCore()
        {
        }

        protected virtual object GetValue()
        {
            return Value.GetValue<object>();
        }

        protected T GetValue<T>()
        {
            object result = GetValue();
            T typedResult = default(T);
            if (result != null)
            {
                typedResult = (T)result;
            }
            return typedResult;
        }

        public abstract void UpdateEditor();

        protected virtual ValidationResult Validate(object value)
        {
            return new ValidationResult()
            {
                IsValid = true,
                Value = value
            };
        }

        protected virtual void SetValue(object value)
        {
            if (guard || value == null)
            {
                return;
            }

            var validation = Validate(value);
            if (!validation.IsValid)
            {
                return;
            }
            value = validation.Value;

            // Avalonia raises a TextBox's TextChanged through the dispatcher, after UpdateEditor
            // has set the text and let go of the guard: without this, showing a figure in the
            // grid recorded one undo step per text row, setting what was already set
            if (Value != null && Equals(Value.GetValue<object>(), value))
            {
                return;
            }

            if (Value != null && Value.CanSetValue)
            {
                guard = true;
                try
                {
                    if (ActionManager != null)
                    {
                        Actions.SetProperty(ActionManager, Value, value);
                    }
                    else
                    {
                        Value.SetValue(value);
                    }
                }
                catch (Exception)
                {
                }
                finally
                {
                    guard = false;
                }
            }
        }
    }
}
