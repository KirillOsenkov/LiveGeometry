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
