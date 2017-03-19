using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;

namespace DynamicGeometry
{
    partial class Behavior
    {
        public static event Action<Behavior> NewBehaviorCreated;
        public static event Action<Behavior> BehaviorDeleted;

        public BehaviorToolButton CreateToolButton()
        {
            return new BehaviorToolButton(this);
        }

        public static IEnumerable<Behavior> LoadBehaviors(Assembly assembly)
        {
            List<Behavior> result = new List<Behavior>();
            Type basic = typeof(Behavior);

            foreach (Type t in assembly.GetTypes())
            {
                if (basic.IsAssignableFrom(t)
                    && !t.IsAbstract
                    && t.GetConstructor(new Type[0]) != null
                    && !t.HasAttribute<IgnoreAttribute>())
                {
                    Behavior instance = Activator.CreateInstance(t) as Behavior;
                    result.Add(instance);
                }
            }

            BehaviorOrderer.Order(result);

            return result;
        }

        protected virtual FreePoint CreatePointAtCurrentPosition(
            Point coordinates)
        {
            var result = Factory.CreateFreePoint(Drawing, coordinates);
            Actions.Add(Drawing, result);
            return result;
        }

        public static void Add(UserDefinedTool newBehavior)
        {
            if (Behavior.NewBehaviorCreated != null)
            {
                Behavior.NewBehaviorCreated(newBehavior);
            }
            ToolStorage.Instance.AddTool(newBehavior);
        }

        public static void Delete(UserDefinedTool tool)
        {
            if (Behavior.BehaviorDeleted != null)
            {
                Behavior.BehaviorDeleted(tool);
            }
            ToolStorage.Instance.RemoveTool(tool);
        }
    }
}
