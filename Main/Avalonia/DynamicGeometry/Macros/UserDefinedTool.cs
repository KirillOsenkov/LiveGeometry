using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Avalonia.Input;

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
            : this(XElement.Parse(macro))
        {
        }

        public UserDefinedTool(XElement macro)
        {
            RootElement = macro;
            mutableName = RootElement.Attribute("Name").Value;
            ReadInputs();
        }

        /// <summary>
        /// A tool from a macro kept between runs; null, with the reason in words, for a text
        /// that is no macro this version can use: damaged, made by a newer version (its
        /// figures may be written in a way this one doesn't read), asking for a kind of
        /// figure this version doesn't have. Checked, not tried: an exception, even caught,
        /// is an error report on screen.
        /// </summary>
        public static UserDefinedTool Read(string text, out string problem)
        {
            problem = null;
            if (string.IsNullOrWhiteSpace(text) || !text.TrimStart().StartsWith("<"))
            {
                problem = "It is not a tool.";
                return null;
            }

            XElement macro;
            try
            {
                macro = XElement.Parse(text);
            }
            catch (System.Xml.XmlException ex)
            {
                problem = "It is damaged: " + ex.Message;
                return null;
            }

            bool hasVersion = double.TryParse(
                macro.ReadString("Version"),
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out double version);
            var inputs = macro.Element("Inputs")?.Elements("Input").ToList();
            if (macro.Name.LocalName != "Macro"
                || macro.ReadString("Name").IsEmpty()
                || inputs == null
                || inputs.Count == 0
                || macro.Element("Figures") == null)
            {
                problem = "It is not a tool.";
            }
            else if (!hasVersion || !double.IsFinite(version))
            {
                problem = "It says no version.";
            }
            else if (version > Settings.CurrentDrawingVersion)
            {
                problem = "It was made by a newer version of Live Geometry.";
            }
            else if (inputs.FirstOrDefault(i => i.ReadString("Name").IsEmpty() || DrawingDeserializer.FindType(i.ReadString("Type") ?? "") == null) is XElement unknown)
            {
                problem = "It starts from a kind of figure this version doesn't have: " + unknown.ReadString("Type") + ".";
            }

            return problem == null ? new UserDefinedTool(macro) : null;
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

            [PropertyGridVisible(false)]
            public UserDefinedTool Tool
            {
                get
                {
                    return Parent;
                }
            }

            // the name on the tool's button: never empty, and no other tool's (the editor says so)
            [PropertyGridVisible]
            [PropertyGridPreferredEditor("ToolName")]
            [PropertyGridEvent("KeyDown", "Name_KeyDown")]
            public string Name
            {
                get
                {
                    return Parent.Name;
                }
                set
                {
                    value = value?.Trim();
                    if (!value.IsEmpty() && !IsToolNameTaken(value, except: Parent))
                    {
                        Parent.MutableName = value;
                    }
                }
            }

            /// <summary>Enter is OK, unless the name was refused (the box says why)</summary>
            public void Name_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.Key == Key.Enter)
                {
                    if (!(sender is StringEditor editor) || string.IsNullOrEmpty(editor.ErrorText))
                    {
                        OK();
                    }

                    e.Handled = true;
                }
            }

            /// <summary>Puts the panel away and keeps the tool: it needs nothing typed to work</summary>
            [PropertyGridVisible]
            [PropertyGridIcon(PropertyGridIcon.Check)]
            public void OK()
            {
                Parent.Drawing?.RaiseDisplayProperties(null);
            }

            [PropertyGridVisible]
            [PropertyGridName("Delete this tool")]
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

        public static UserDefinedTool AddFromString(string macro)
        {
            UserDefinedTool tool = new UserDefinedTool(macro);
            Behavior.Add(tool);
            return tool;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            // a copy: the styles brought along may be renamed in it, and the macro keeps the names
            var figuresElement = new XElement(RootElement.Element("Figures"));
            var inputs = new Dictionary<string, IFigure>();
            for (int i = 0; i < Inputs.Count; i++)
            {
                inputs.Add(Inputs[i].Name, FoundDependencies[i]);
            }
            var deserializer = new DrawingDeserializer();
            var styles = RootElement.Element("Styles");
            if (styles != null)
            {
                // as a paste does; not recorded (nor is Create new style): a style left behind
                // by undo is not saved, files carry only the styles their figures name
                PasteAction.BringStyles(Drawing, styles, figuresElement, deserializer);
            }

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

            // A figure that had its default name in the macro (segment AB) gets the default
            // name here (CD, after the points it was given; an empty name is the default).
            // It kept AB, which on C and D reads like a typed name, unless another figure in
            // the drawing had it. After the expressions are bound: the rename rewrites the
            // texts that name it.
            var defaultNamed = RootElement.Element("Figures").Elements()
                .Where(e => e.ReadBool("DefaultName", defaultValue: false))
                .Select(e => e.ReadString("Name"))
                .ToList();
            foreach (var name in defaultNamed)
            {
                if (byMacroName.TryGetValue(name, out var figure) && figures.Contains(figure))
                {
                    figure.Name = "";
                }
            }
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