using System;
using System.Net;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Animation;
using Avalonia.Controls.Shapes;

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
