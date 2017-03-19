using System.Collections.Generic;
using System.ComponentModel;

namespace DynamicGeometry
{
    public class BehaviorOrderer
    {
        public static void Order(List<Behavior> behaviors)
        {
            behaviors.Sort(Compare);
        }

        public static string GetCategory(Behavior behavior)
        {
            if (behavior is UserDefinedTool)
            {
                return BehaviorCategories.Custom;
            }

            var result = BehaviorCategories.Misc;
            var categoryAttribute = behavior.GetType().GetAttribute<CategoryAttribute>();
            if (categoryAttribute != null)
            {
                result = categoryAttribute.Category;
            }
            return result;
        }

        public static int Compare(Behavior behavior1, Behavior behavior2)
        {
            var orderAttribute1 = behavior1.GetType().GetAttribute<OrderAttribute>();
            var categoryAttribute1 = behavior1.GetType().GetAttribute<CategoryAttribute>();
            var orderAttribute2 = behavior2.GetType().GetAttribute<OrderAttribute>();
            var categoryAttribute2 = behavior2.GetType().GetAttribute<CategoryAttribute>();

            if (orderAttribute2 == null || categoryAttribute2 == null)
            {
                return 1;
            }

            if (orderAttribute1 == null || categoryAttribute1 == null)
            {
                return -1;
            }

            var categoryIndex1 = GetCategoryIndex(categoryAttribute1.Category);
            var categoryIndex2 = GetCategoryIndex(categoryAttribute2.Category);

            if (categoryIndex1 == categoryIndex2)
            {
                return orderAttribute1.Order.CompareTo(orderAttribute2.Order);
            }

            return categoryIndex1.CompareTo(categoryIndex2);
        }

        private static int GetCategoryIndex(string categoryName)
        {
            int result = int.MaxValue;
            CategoryOrder.TryGetValue(categoryName, out result);
            return result;
        }

        private static Dictionary<string, int> categoryOrder;
        private static Dictionary<string, int> CategoryOrder
        {
            get
            {
                if (categoryOrder == null)
                {
                    categoryOrder = new Dictionary<string, int>();
                    var fields = typeof(BehaviorCategories).GetFields();
                    int index = 0;
                    foreach (var field in fields)
                    {
                        categoryOrder.Add(field.Name, index++);
                    }
                }
                return categoryOrder;
            }
        }
    }
}
