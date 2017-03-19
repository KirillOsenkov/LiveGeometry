using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public class ToolStorage
    {
        static ToolStorage instance;
        public static ToolStorage Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ToolStorage();
                }
                return instance;
            }
            set
            {
                instance = value;
            }
        }

        public virtual void AddTool(UserDefinedTool newBehavior)
        {

        }

        public virtual void RenameTool(UserDefinedTool behavior, string newName)
        {

        }

        public virtual void RemoveTool(UserDefinedTool behavior)
        {

        }
    }
}
