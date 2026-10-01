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

        /// <summary>
        /// Whether the next set of the same property may join this one in a single undo step:
        /// the letters typed into a box, the positions a slider or a color picker passes
        /// through. A check box or a choice from a list is a step each time.
        /// </summary>
        public bool Coalesce { get; set; }

        /// <summary>
        /// The run of edits the set belongs to, when its editor tells them apart: only sets
        /// of one run join. An editor starts a new run when it is done with a value (Enter,
        /// leaving the box) and every time the grid shows the object anew - otherwise a
        /// length typed now joined the one typed five minutes ago, if nothing else was done
        /// in between, and undo skipped it.
        /// </summary>
        public object Run { get; set; }

        /// <summary>Whose history the action is in, to know when there is something to redo</summary>
        public ActionManager ActionManager { get; set; }

        // one value per object of a multiple selection, each with what it had before: the
        // objects may have differed, and undo gives each its own back
        IValueProvider[] targets;
        object[] oldStates;

        static IValueProvider[] GetTargets(IValueProvider property)
        {
            return property is CompositeValueProvider composite
                ? composite.InnerList.ToArray()
                : new[] { property };
        }

        protected override void ExecuteCore()
        {
            targets = GetTargets(Property);
            oldStates = new object[targets.Length];
            for (int i = 0; i < targets.Length; i++)
            {
                oldStates[i] = targets[i] is IRestorableValue restorable
                    ? restorable.CaptureState()
                    : targets[i].GetValue<object>();
            }

            Property.SetValue(NewValue);
        }

        protected override void UnExecuteCore()
        {
            // last to first: a set that takes a figure out of the drawing (a point's label)
            // remembers where it was, and those places only add up in reverse
            for (int i = targets.Length - 1; i >= 0; i--)
            {
                if (targets[i] is IRestorableValue restorable)
                {
                    restorable.RestoreState(oldStates[i]);
                }
                else
                {
                    targets[i].SetValue(oldStates[i]);
                }
            }
        }

        public override bool TryToMerge(IAction followingAction)
        {
            SetPropertyAction next = followingAction as SetPropertyAction;
            if (next == null || !Coalesce || !next.Coalesce || targets == null || !Equals(Run, next.Run))
            {
                return false;
            }

            // the history keeps what there is to redo when an action is merged into the last
            // one: a new edit must end that, so it is an action of its own
            if (ActionManager != null && ActionManager.CanRedo)
            {
                return false;
            }

            var nextTargets = GetTargets(next.Property);
            if (nextTargets.Length != targets.Length)
            {
                return false;
            }

            for (int i = 0; i < targets.Length; i++)
            {
                if (!IsSameValue(targets[i], nextTargets[i]))
                {
                    return false;
                }
            }

            this.NewValue = next.NewValue;
            Property.SetValue(NewValue);
            return true;
        }

        /// <summary>The same property of the same object, seen under the same theme</summary>
        static bool IsSameValue(IValueProvider first, IValueProvider second)
        {
            return first.GetType() == second.GetType()
                && first.Name == second.Name
                && first.Parent != null
                && first.Parent == second.Parent
                && first.CanSetValue == second.CanSetValue
                && (first as ThemedValue)?.Theme == (second as ThemedValue)?.Theme;
        }
    }
}
