using GuiLabs.Undo;

namespace DynamicGeometry
{
    public class SetPropertyAction : AbstractAction
    {
        public SetPropertyAction(
            IValueProvider property, object newValue)
        {
            Property = property;
            NewValue = newValue;
        }

        public IValueProvider Property { get; set; }
        public object NewValue { get; set; }
        public object OldValue { get; set; }

        protected override void ExecuteCore()
        {
            OldValue = Property.GetValue<object>();
            Property.SetValue(NewValue);
        }

        protected override void UnExecuteCore()
        {
            Property.SetValue(OldValue);
        }

        public override bool TryToMerge(IAction followingAction)
        {
            SetPropertyAction next = followingAction as SetPropertyAction;

            // Comparing the Property does not allow for proper merging. - D.H.
            // if (next != null
            // && next.Property == this.Property)
            // Using new comparison.
            if (next != null
                && next.Property.Name == this.Property.Name
                && next.Property.Parent == this.Property.Parent
                && next.Property.CanSetValue == this.Property.CanSetValue)
            {
                this.NewValue = next.NewValue;
                Property.SetValue(NewValue);
                return true;
            }
            return false;
        }
    }
}
