using System.Reflection;

namespace DynamicGeometry
{
    public abstract partial class BaseValueEditorFactory<T> : IValueEditorFactory
        where T : IValueEditor, new()
    {
        public virtual bool SupportsValue(IValueProvider property)
        {
            return true;
        }

        public virtual IValueEditor CreateEditor(IValueProvider property)
        {
            var editor = new T() { Value = property };
            return editor;
        }

        public int LoadOrder { get; set; }
    }

    public abstract partial class BaseValueEditorFactory<TProperty, TDataType> :
        BaseValueEditorFactory<TProperty>
        where TProperty : IValueEditor, new()
    {
        public override bool SupportsValue(IValueProvider value)
        {
            return value.Type == typeof(TDataType) && base.SupportsValue(value);
        }
    }
}
