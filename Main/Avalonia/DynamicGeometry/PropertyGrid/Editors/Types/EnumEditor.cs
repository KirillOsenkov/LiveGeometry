using System;
using System.Linq;
using System.Reflection;

namespace DynamicGeometry
{
    public class EnumEditorFactory : BaseValueEditorFactory<EnumEditor>
    {
        public override bool SupportsValue(IValueProvider value)
        {
            return value.Type.IsEnum && base.SupportsValue(value);
        }
    }

    public class EnumEditor : SelectorValueEditor, IValueEditor
    {
        public override void FillList()
        {
            var propertyType = Value.Type;
            Items = from f in propertyType.GetFields()
                    where f.FieldType == propertyType
                    select ShownName(Enum.Parse(propertyType, f.Name, true));
            base.FillList();
        }

        /// <summary>What the list says for a value: its [PropertyGridName] ("Catmull-Rom"), else its name</summary>
        string ShownName(object value)
        {
            var name = value.ToString();
            var field = Value.Type.GetField(name, BindingFlags.Public | BindingFlags.Static);
            return field?.GetAttribute<PropertyGridNameAttribute>()?.Name ?? name;
        }

        protected override ValidationResult Validate(object value)
        {
            var text = value.ToString();
            var field = Value.Type
                .GetFields(BindingFlags.Public | BindingFlags.Static)
                .FirstOrDefault(f => f.GetAttribute<PropertyGridNameAttribute>()?.Name == text);
            return new ValidationResult()
            {
                IsValid = true,
                Value = field != null ? field.GetValue(null) : Enum.Parse(Value.Type, text, false)
            };
        }

        public override void UpdateEditor()
        {
            // no value: several figures that differ (it threw on the null, with everything
            // selected in a drawing whose measurements have different units)
            var value = GetValue();
            var shown = value != null ? ShownName(value) : null;
            ShowSelected(item => item.Equals(shown));
        }
    }
}
