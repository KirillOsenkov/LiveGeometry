using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Avalonia;

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
            tools.AddRange(result);
            return result;
        }

        // the tools on the ribbon: the library's and those the user defined
        static readonly List<Behavior> tools = new List<Behavior>();

        /// <summary>
        /// The ribbon's instance of a tool, for code that picks the tool itself (Edit figures
        /// of a show/hide box): the ribbon finds its button by the instance. Null when the
        /// tools were never loaded (a drawing without the editor).
        /// </summary>
        public static T FindTool<T>() where T : Behavior
        {
            return tools.OfType<T>().FirstOrDefault();
        }

        /// <summary>Whether a tool on the ribbon other than <paramref name="except"/> is called that, in any case</summary>
        public static bool IsToolNameTaken(string name, Behavior except = null)
        {
            return tools.Exists(t => t != except && string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>The name, or with a number after it when a tool has it already: Catenary 2</summary>
        public static string UniqueToolName(string name)
        {
            var candidate = name;
            for (int i = 2; IsToolNameTaken(candidate); i++)
            {
                candidate = name + " " + i;
            }

            return candidate;
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
            tools.Add(newBehavior);
            if (Behavior.NewBehaviorCreated != null)
            {
                Behavior.NewBehaviorCreated(newBehavior);
            }
            ToolStorage.Instance.AddTool(newBehavior);
        }

        public static void Delete(UserDefinedTool tool)
        {
            tools.Remove(tool);
            if (Behavior.BehaviorDeleted != null)
            {
                Behavior.BehaviorDeleted(tool);
            }
            ToolStorage.Instance.RemoveTool(tool);
        }
    }
}
