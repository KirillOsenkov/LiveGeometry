using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

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
        [PropertyGridNoUndo]
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
            [PropertyGridDestructive]
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
            var tempFigures = deserializer.ReadFigures(figuresElement, Drawing, inputs).ToList();
            byMacroName = inputs;
            return tempFigures;
        }

        // the inputs and the figures made, by the names the macro gives them
        Dictionary<string, IFigure> byMacroName;

        protected override void FiguresAdded(IList<IFigure> figures)
        {
            RebindExpressions(figures);
        }

        protected override void CreateTempResults()
        {
            base.CreateTempResults();
            RebindExpressions(TempResults);
        }

        /// <summary>
        /// The expressions of the figures made (a point by coordinates, a label's [AB]) name
        /// the inputs and the other figures made, by the names the macro gives them. Compiled
        /// as they are read, they named the figures the tool was defined on: a catenary made
        /// from two other points was the first one again, only the segment between the new
        /// points and the point sliding on it being new. As a paste does
        /// (<see cref="PasteAction"/>): the texts are rewritten to the names the figures have
        /// in the drawing and compiled again, once they are in.
        /// </summary>
        void RebindExpressions(IEnumerable<IFigure> figures)
        {
            if (byMacroName == null)
            {
                return;
            }

            var oldNames = new Dictionary<IFigure, string>();
            foreach (var pair in byMacroName)
            {
                oldNames[pair.Value] = pair.Key;
            }

            var renamer = new ExpressionRenamer(Drawing, oldNames, preferred: byMacroName.Values.ToArray());
            var holders = figures.OfType<IRenamableExpressions>().ToArray();
            foreach (var holder in holders)
            {
                holder.RenameInExpressions(renamer);
            }

            foreach (var holder in holders)
            {
                holder.RebindExpressions();
            }
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

        /// <summary>
        /// What to click, in order, with what each stood for when the tool was defined:
        /// "Click a slider (rope), a point (A), then a point (B)." A click on another kind
        /// than the one asked for makes a point or nothing, so the order must be known.
        /// </summary>
        public override string HintText
        {
            get
            {
                var steps = Inputs.Select(Describe).ToList();
                if (steps.Count > 1)
                {
                    steps[steps.Count - 1] = "then " + steps[steps.Count - 1];
                }

                return "Click " + string.Join(", ", steps) + ".";
            }
        }

        public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
        {
            // the point following the cursor is among the found ones already
            int next = FoundDependencies.Count - (TempPoint != null ? 1 : 0);
            if (next < 0 || next >= Inputs.Count)
            {
                return base.ConstructionHintText(args);
            }

            return "Click " + Describe(Inputs[next]) + ".";
        }

        static string Describe(InputInfo input)
        {
            return DescribeFigureType(input.Type) + " (" + input.Name + ")";
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
            return MacroIcon.Build(RootElement);
        }

    }
}