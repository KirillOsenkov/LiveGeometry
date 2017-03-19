using System.Windows.Controls;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class FontFamilyEditorFactory : BaseValueEditorFactory<FontFamilyEditor>
    {
        public override bool SupportsValue(IValueProvider value)
        {
            return value.Type == typeof(FontFamily) && base.SupportsValue(value);
        }
    }

    public class FontFamilyEditor : SelectorValueEditor, IValueEditor
    {
        public override void FillList()
        {
            Items = new[]
            {
                "Arial",
                "Arial Black",
                "Arial Unicode MS",
                "Calibri",
                "Cambria",
                "Cambria Math",
                "Comic Sans MS",
                "Candara",
                "Consolas",
                "Constantia",
                "Corbel",
                "Courier New",
                "Georgia",
                "Lucida Sans Unicode",
                "Segoe UI",
                "Symbol",
                "Tahoma",
                "Times New Roman",
                "Trebuchet MS",
                "Verdana",
                "Wingdings",
                "Wingdings 2",
                "Wingdings 3",
            };
            base.FillList();
        }

        protected override ValidationResult Validate(object value)
        {
            return new ValidationResult()
            {
                IsValid = true,
                Value = new FontFamily(value.ToString())
            };
        }

        public override void UpdateEditor()
        {
            var value = GetValue();
            foreach (var item in Items)
            {
                if (item.Equals(value.ToString()))
                {
                    Selector.SelectedItem = item;
                }
            }
        }
    }
}
