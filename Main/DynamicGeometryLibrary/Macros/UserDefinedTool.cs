using System;
using System.Collections.Generic;
using System.Windows.Input;
using System.Xml.Linq;
using System.Windows;

namespace DynamicGeometry
{
    public class InputInfo
    {
        public string Name { get; set; }
        public Type Type { get; set; }
    }

    [Ignore]
    public class UserDefinedTool : FigureCreator
    {
        public UserDefinedTool()
        {
        }

        public UserDefinedTool(string macro)
        {
            RootElement = XElement.Parse(macro);
            mutableName = RootElement.Attribute("Name").Value; 
            ReadInputs();
        }

        [PropertyGridName("Tool properties")]
        public class UserDefinedDialog
        {
            public UserDefinedDialog(UserDefinedTool parent)
            {
                Parent = parent;
            }

            //[PropertyGridVisible]
            public string XML
            {
                get
                {
                    return Parent.RootElement.ToString();
                }
            }

            [PropertyGridVisible]
            public string Name
            {
                get
                {
                    return Parent.Name;
                }
                set
                {
                    Parent.MutableName = value;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Delete this tool button")]
            public void Delete()
            {
                Parent.AbortAndSetDefaultTool();
                Behavior.Delete(Parent);
            }

            UserDefinedTool Parent;
        }

        UserDefinedDialog dialog;

        public override object PropertyBag
        {
            get
            {
                if (dialog == null)
                {
                    dialog = new UserDefinedDialog(this);
                }
                return dialog;
            }
        }

        private void ReadInputs()
        {
            var inputs = RootElement.Element("Inputs");
            foreach (var inputElement in inputs.Elements())
            {
                string name = inputElement.Attribute("Name").Value;
                string typeName = inputElement.Attribute("Type").Value;
                Type type = DrawingDeserializer.FindType(typeName);
                Inputs.Add(new InputInfo()
                {
                    Name = name,
                    Type = type
                });
            }
        }

        List<InputInfo> Inputs = new List<InputInfo>();

        public XElement RootElement { get; set; }

        public override string Name
        {
            get { return MutableName; }
        }

        string mutableName = "Custom tool";
        public string MutableName
        {
            get
            {
                return mutableName;
            }
            set
            {
                mutableName = value;
                var attribute = RootElement.Attribute("Name");
                if (attribute == null)
                {
                    attribute = new XAttribute("Name", value);
                    RootElement.Add(attribute);
                }
                else
                {
                    attribute.Value = value;
                }
                RaisePropertyChanged("Name");
                ToolStorage.Instance.RenameTool(this, value);
            }
        }

        public static void AddFromString(string macro)
        {
            UserDefinedTool tool = new UserDefinedTool(macro);
            Behavior.Add(tool);
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            var figuresElement = RootElement.Element("Figures");
            var inputs = new Dictionary<string, IFigure>();
            for (int i = 0; i < Inputs.Count; i++)
            {
                inputs.Add(Inputs[i].Name, FoundDependencies[i]);
            }
            var deserializer = new DrawingDeserializer();
            //EnsureUniqueNames(Drawing, figuresElement);   This changes RootElement so that the names don't match up with Inputs. - D.H.
            var tempFigures = deserializer.ReadFigures(figuresElement, Drawing, inputs);
            return tempFigures;
        }

        private void EnsureUniqueNames(Drawing drawing, XElement figuresElement)
        {
            List<XElement> renames = new List<XElement>();

            foreach (var figureElement in figuresElement.Elements())
            {
                string oldName = figureElement.ReadString("Name");
                if (drawing.Figures[oldName] != null)
                {
                    renames.Add(figureElement);
                }
            }

            if (renames.Count == 0)
            {
                return;
            }

            foreach (var element in renames)
            {
                string oldName = element.ReadString("Name");
                string newName = GenerateUniqueName(drawing, oldName);
                element.SetAttributeValue("Name", newName);
                foreach (var figure in figuresElement.Elements())
                {
                    foreach (var dependency in figure.Elements("Dependency"))
                    {
                        if (dependency.ReadString("Name") == oldName)
                        {
                            dependency.SetAttributeValue("Name", newName);
                        }
                    }
                }
            }
        }

        private string GenerateUniqueName(Drawing drawing, string originalName)
        {
            while (drawing.Figures[originalName] != null)
            {
                originalName += "1";
            }

            return originalName;
        }

        protected override DependencyList InitExpectedDependencies()
        {
            DependencyList result = new DependencyList();
            foreach (var input in Inputs)
            {
                result.Add(input.Type);
            }
            return result;
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            Drawing.RaiseSelectionChanged();
            base.MouseDown(sender, e);
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Point(0.5, 0.5)
                .Canvas;
        }

    }
}