using System;
using System.Linq;

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
                    select Enum.Parse(propertyType, f.Name, true).ToString();
            base.FillList();
        }

        protected override ValidationResult Validate(object value)
        {
            return new ValidationResult()
            {
                IsValid = true,
                Value = Enum.Parse(Value.Type, value.ToString(), false)
            };
        }

        public override void UpdateEditor()
        {
            // no value: several figures that differ (it threw on the null, with everything
            // selected in a drawing whose measurements have different units)
            var value = GetValue()?.ToString();
            ShowSelected(item => item.Equals(value));
        }
    }
}
