using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using GuiLabs.Undo;
using System.Xml;
using System.Xml.Linq;
using System.Linq;

namespace DynamicGeometry
{
    public abstract partial class FigureBase :
        IFigure,
        IPropertyGridHost,
        IPropertyGridContentProvider,
        INotifyPropertyChanged,
        INotifyPropertyChanging
    {
        public FigureBase()
        {
            Exists = true;
            IsHitTestVisible = true;
            mDependencies.CollectionChanged += mDependencies_CollectionChanged;
        }

        public virtual string GenerateFigureName()
        {
            return GenerateFigureName(null);
        }

        public virtual string GenerateFigureName(List<string> blacklist)
        {
            // a composite's parts are numbered: only the composite takes the name of its points
            bool inDrawing = Drawing != null && Drawing.Figures.Contains(this);
            var stem = inDrawing ? NameFromDependencies() : null;
            if (stem == null)
            {
                // Hidden helpers too (the ellipse's short axis), unlike hidden points: other
                // figures' constructions name them ("of line g and segment AB"), and a name
                // like PerpendicularLine1 was one the list showed nowhere else. Not a part of
                // a composite (a vector's shaft). A copy being pasted isn't in the list yet,
                // and takes one.
                bool isPart = Drawing != null && Drawing.Figures.FindTopLevel(this) is IFigure whole && whole != this;
                if (Drawing != null && !isPart && FirstLetter != null)
                {
                    return GenerateLetterName(this, FirstLetter, blacklist);
                }

                return this.GenerateNewName();
            }

            // segment AB and line AB side by side: AB and AB2
            for (int i = 1; ; i++)
            {
                var candidate = i == 1 ? stem : stem + i;
                if (this.NameAvailable(candidate) && (blacklist == null || !blacklist.Contains(candidate)))
                {
                    return candidate;
                }
            }
        }

        /// <summary>
        /// The lowercase letters a figure not named after its points is named with, as a
        /// textbook names line g or circle c and as a slider is named: not e (the constant),
        /// not x and y (the axes and the variable of a function), not the letters that read
        /// as digits
        /// </summary>
        public const string Letters = "abcdfghkmnpqrstuvwz";

        /// <summary>
        /// The letter of <see cref="Letters"/> the names of the figure's kind start from
        /// (g for a line, c for a circle), when nothing names it after its points; null for
        /// a figure numbered by its type (Label1, the helpers a measurement or a mark is)
        /// </summary>
        protected virtual string FirstLetter
        {
            get
            {
                return null;
            }
        }

        /// <summary>Whether the figure's kind is named with letters (line g), not numbered (Label1)</summary>
        public bool IsNamedWithLetters
        {
            get
            {
                return FirstLetter != null;
            }
        }

        /// <summary>
        /// The first free name from <paramref name="firstLetter"/> to the end of <see cref="Letters"/>,
        /// then the same with 1, 2... after it. Not round to the start: a function would be
        /// named a, a slider's letter, once lines and circles had taken f to z.
        /// </summary>
        public static string GenerateLetterName(IFigure figure, string firstLetter, List<string> blacklist)
        {
            int start = System.Math.Max(0, Letters.IndexOf(firstLetter, StringComparison.Ordinal));
            for (int i = 0; ; i++)
            {
                for (int j = start; j < Letters.Length; j++)
                {
                    var letter = Letters[j];
                    var candidate = i == 0 ? letter.ToString() : letter + i.ToString();
                    if (figure.NameAvailable(candidate) && (blacklist == null || !blacklist.Contains(candidate)))
                    {
                        return candidate;
                    }
                }
            }
        }

        /// <summary>
        /// The names the figure can take from what it is built on - segment AB or BA, triangle
        /// ABC or BCA... - the one it takes first; null for a figure numbered by its type (Circle1)
        /// </summary>
        protected virtual IReadOnlyList<string> NamesFromDependencies()
        {
            return null;
        }

        string NameFromDependencies()
        {
            var names = NamesFromDependencies();
            return names != null && names.Count > 0 ? names[0] : null;
        }

        /// <summary>How the points of a figure may be read and still name it</summary>
        protected enum PointOrder
        {
            /// <summary>As they are: ray AB is not ray BA</summary>
            Fixed,

            /// <summary>Either way: segment BA is segment AB</summary>
            Reversible,

            /// <summary>From any vertex, either way round: triangle CBA is triangle ABC</summary>
            Cyclic
        }

        /// <summary>
        /// The names of the dependencies run together, if they are all named points, in every
        /// order that names the same figure; the one that comes first alphabetically (in the
        /// order points are named) first: ABCE rather than ECBA.
        /// </summary>
        protected IReadOnlyList<string> NamesFromPoints(PointOrder order, int maxCount = 2)
        {
            return NamesFromPoints(Dependencies, order, maxCount);
        }

        /// <summary>The same for some of the dependencies (a Bezier path's anchors, not its handles)</summary>
        protected IReadOnlyList<string> NamesFromPoints(IList<IFigure> points, PointOrder order, int maxCount)
        {
            if (points.Count == 0 || points.Count > maxCount)
            {
                return null;
            }

            var names = new string[points.Count];
            for (int i = 0; i < names.Length; i++)
            {
                if (!(points[i] is IPoint) || string.IsNullOrEmpty(points[i].Name))
                {
                    return null;
                }

                names[i] = points[i].Name;
            }

            var readings = new List<string[]>() { names };
            if (order != PointOrder.Fixed)
            {
                readings.Add(Enumerable.Reverse(names).ToArray());
            }

            if (order == PointOrder.Cyclic)
            {
                foreach (var reading in readings.ToArray())
                {
                    for (int start = 1; start < reading.Length; start++)
                    {
                        readings.Add(reading.Skip(start).Concat(reading.Take(start)).ToArray());
                    }
                }
            }

            readings.Sort(CompareReadings);
            return readings.Select(reading => string.Concat(reading)).Distinct().ToList();
        }

        static int CompareReadings(string[] first, string[] second)
        {
            for (int i = 0; i < first.Length; i++)
            {
                int result = ComparePointNames(first[i], second[i]);
                if (result != 0)
                {
                    return result;
                }
            }

            return 0;
        }

        /// <summary>In the order points get their names: A to Z, then A1 to Z1</summary>
        static int ComparePointNames(string first, string second)
        {
            var (firstLetters, firstNumber) = SplitNumber(first);
            var (secondLetters, secondNumber) = SplitNumber(second);
            int result = firstNumber.CompareTo(secondNumber);
            return result != 0 ? result : string.CompareOrdinal(firstLetters, secondLetters);
        }

        static (string Letters, long Number) SplitNumber(string name)
        {
            int digits = name.Length;
            while (digits > 0 && char.IsDigit(name[digits - 1]))
            {
                digits--;
            }

            long.TryParse(name.Substring(digits), out long number);
            return (name.Substring(0, digits), number);
        }

        /// <summary>
        /// Nobody has named the figure: it is called what the library calls it (AB after its
        /// points, Circle1 after its type), and a name taken from its points follows them when
        /// one of them is renamed or replaced. Not stored in files: a name that reads like the
        /// default is the default (<see cref="IsDefaultName"/>).
        /// </summary>
        public bool HasDefaultName { get; private set; }

        /// <summary>
        /// The name from the points in any order that names the figure (AB, BA, or AB2 when AB
        /// was taken) or from the type (Segment1, which every figure was called until 2026-09-27)
        /// </summary>
        bool IsDefaultName(string name)
        {
            var names = NamesFromDependencies();
            return names != null && names.Any(stem => IsStemAndNumber(name, stem))
                || IsStemAndNumber(name, GetType().Name);
        }

        static bool IsStemAndNumber(string name, string stem)
        {
            if (stem == null || !name.StartsWith(stem, StringComparison.Ordinal))
            {
                return false;
            }

            for (int i = stem.Length; i < name.Length; i++)
            {
                if (!char.IsDigit(name[i]))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Whether a name typed for the figure would stay. One that reads like the name of
        /// its points with another number, or the other way round (AB3 or BA for segment
        /// AB), is a default name, and those are worked out: it would be put right at
        /// once, leaving an undo step that undoes nothing. The exception is the default
        /// name itself, for a figure that has a name of its own and is to lose it.
        /// </summary>
        public bool KeepsTypedName(string name)
        {
            if (name == mName || NameFromDependencies() == null || !IsDefaultName(name))
            {
                return true;
            }

            return !HasDefaultName && name == GenerateFigureName(null);
        }

        /// <summary>A default name follows the points the figure is built on</summary>
        public void UpdateDefaultName()
        {
            if (!HasDefaultName
                || Drawing == null
                || NameFromDependencies() == null
                || !Drawing.Figures.Contains(this))
            {
                return;
            }

            var name = GenerateFigureName(null);
            if (name != mName)
            {
                Name = name;
            }
        }

        #region Numbers of default names

        [ThreadStatic]
        static bool settlingDefaultNames;

        /// <summary>
        /// Figures named after the same points - segment AB, line AB, a second segment AB -
        /// are AB, AB2, AB3 in the order they have in the drawing's list: the numbers are
        /// worked out from the drawing as it is, not from what happened to it. They used to
        /// go to whoever asked first, so AB2 stayed AB2 after AB was deleted until the file
        /// was opened again, and undo of a join handed the two names back the other way
        /// round. Called for a figure that came into the list, moved in it or was renamed;
        /// puts the numbers of its whole group right.
        /// </summary>
        public static void SettleDefaultNames(Drawing drawing, IFigure figure)
        {
            if (drawing == null || settlingDefaultNames || drawing.IsReading)
            {
                return;
            }

            var stem = (figure as FigureBase)?.NameFromDependencies();
            if (stem == null)
            {
                return;
            }

            // the others named after the same points are built on the same points
            var figures = drawing.Figures;
            var group = new List<FigureBase>();
            foreach (var dependent in figure.Dependencies[0].Dependents)
            {
                if (dependent is FigureBase candidate
                    && candidate.HasDefaultName
                    && !group.Contains(candidate)
                    && figures.Contains(candidate)
                    && candidate.NameFromDependencies() == stem)
                {
                    group.Add(candidate);
                }
            }

            if (group.Count == 0)
            {
                return;
            }

            group.Sort((first, second) => figures.IndexOf(first).CompareTo(figures.IndexOf(second)));
            settlingDefaultNames = true;
            try
            {
                int number = 0;
                foreach (var member in group)
                {
                    // AB, AB2, AB3... but for a name something else in the drawing has
                    string name;
                    do
                    {
                        number++;
                        name = number == 1 ? stem : stem + number;
                    }
                    while (figures.Any(f => f.Name == name && !group.Contains(f as FigureBase)));

                    // (whichever of the group has the name gives it up: the setter sees to that)
                    if (member.mName != name)
                    {
                        member.Name = name;
                    }
                }
            }
            finally
            {
                settlingDefaultNames = false;
            }
        }

        /// <summary>
        /// A name has become free - its figure was deleted or renamed: the figures numbered
        /// after it move up (AB2 is AB once AB is gone)
        /// </summary>
        public static void SettleDefaultNamesAfter(Drawing drawing, string freedName)
        {
            if (drawing == null || settlingDefaultNames || drawing.IsReading || string.IsNullOrEmpty(freedName))
            {
                return;
            }

            // one figure of each group the name could go to (a group shares its first letter)
            var stems = new Dictionary<string, FigureBase>();
            foreach (var figure in drawing.Figures)
            {
                if (figure is FigureBase candidate
                    && candidate.HasDefaultName
                    && !string.IsNullOrEmpty(candidate.mName)
                    && candidate.mName[0] == freedName[0])
                {
                    var stem = candidate.NameFromDependencies();
                    if (stem != null && !stems.ContainsKey(stem) && IsStemAndNumber(freedName, stem))
                    {
                        stems.Add(stem, candidate);
                    }
                }
            }

            foreach (var member in stems.Values)
            {
                SettleDefaultNames(drawing, member);
            }
        }

        /// <summary>After a wave of renames: the groups the renamed figures are in now, and the ones their old names were in</summary>
        static void SettleDefaultNames(Dictionary<IFigure, string> wave)
        {
            if (settlingDefaultNames)
            {
                return;
            }

            foreach (var pair in wave)
            {
                var drawing = pair.Key.Drawing;
                if (drawing == null || pair.Key.Name == pair.Value || !drawing.Figures.Contains(pair.Key))
                {
                    continue;
                }

                SettleDefaultNames(drawing, pair.Key);
                SettleDefaultNamesAfter(drawing, pair.Value);
            }
        }

        #endregion

        protected Drawing drawing;
        public virtual Drawing Drawing
        {
            get
            {
                return drawing;
            }
            set
            {
                drawing = value;
            }
        }

        public virtual void OnAddingToDrawing(Drawing drawing)
        {
            this.GenerateNewNameIfNecessary(drawing, null);
        }

        /// <summary>The label writing the figure's name next to it, while it shows one (<see cref="FigureLabel"/>)</summary>
        public FigureLabel NameLabel { get; set; }

        // the label the figure had last: showing the name again brings the same one back,
        // where it was dragged to, and what the undo history did to it still applies to it
        FigureLabel retiredNameLabel;

        /// <summary>
        /// Whether the figure writes its name next to itself. Setting it adds or removes the
        /// label directly, like a point's ShowName: the property set is the undo step. The
        /// label taken away is kept, and comes back when the name is shown again (undo
        /// included). The figures that offer it in the grid (lines, circles) expose it as
        /// "Show name".
        /// </summary>
        public bool HasNameLabel
        {
            get
            {
                return NameLabel != null;
            }
            set
            {
                if (value == (NameLabel != null) || Drawing == null)
                {
                    return;
                }

                if (value)
                {
                    NameLabel = retiredNameLabel ?? Factory.CreateFigureLabel(Drawing, this);
                    NameLabel.Visible = true;
                    retiredNameLabel = null;
                    Drawing.Figures.Return(NameLabel, owner: this);
                }
                else
                {
                    var label = NameLabel;
                    NameLabel = null;
                    Drawing.Figures.Retire(label);
                    retiredNameLabel = label;
                }
            }
        }

        public virtual void OnRemovingFromDrawing(Drawing drawing)
        {
        }

        //public static int ID { get; set; } - Phased out 8/11/2011. D.H.

        public virtual IFigure Clone()
        {
            // Updated: clone now inherits properties using read/writeXml. - D.H.
            var newFigure = Activator.CreateInstance(this.GetType()) as IFigure;
            newFigure.Dependencies.AddRange(this.Dependencies);
            newFigure.RegisterWithDependencies();
            var s = new System.Text.StringBuilder();
            using (var w = System.Xml.XmlWriter.Create(s, new System.Xml.XmlWriterSettings()))
            {
                w.WriteStartElement(this.GetType().Name);
                WriteXml(w);
                w.WriteEndElement();
            }
            var xml = s.ToString();
            newFigure.Drawing = Drawing;
            try
            {
                newFigure.ReadXml(XElement.Parse(xml));
            }
            catch (Exception ex)
            {
                Drawing.RaiseError(this, ex);
            }
            return newFigure;
        }

        public override string ToString()
        {
            return Name;
        }

        /// <summary>
        /// What the figure is, in a word or two ("Segment", "Triangle"), for <see cref="Title"/>;
        /// null where no word fits
        /// </summary>
        protected virtual string Kind
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// "Segment AB", "Triangle ABC": the kind in front of the name, for the property grid -
        /// unless the name says it already (Circle1, Bezier3), or the figure goes by what it
        /// is built on and has a number for a name (Distance, not DistanceMeasurement1:
        /// <see cref="NamedByConstruction"/>)
        /// </summary>
        public string Title
        {
            get
            {
                var (kind, name) = TitleParts;
                return kind == null ? name : name == null ? kind : kind + " " + name;
            }
        }

        /// <summary>
        /// The two parts of the <see cref="Title"/>, for whoever draws them apart (the Figure
        /// List): the kind ("Midpoint") and the name as shown ("E", "A₁"); either may be null
        /// </summary>
        public (string Kind, string Name) TitleParts
        {
            get
            {
                // without a kind, what the figure says of itself ("Coordinate grid")
                var kind = Kind;
                if (kind == null || string.IsNullOrEmpty(Name))
                {
                    return (null, ToString());
                }

                // Distance, not DistanceMeasurement1. Only for what nothing refers to by name:
                // Circle1 stays, since "on Circle₁" in another row must be found in the list.
                if (NamedByConstruction && IsStemAndNumber(Name, GetType().Name))
                {
                    return (kind, null);
                }

                if (NameSays(kind))
                {
                    return (null, NameDisplay.Format(Name));
                }

                return (kind, NameDisplay.Format(Name));
            }
        }

        /// <summary>Whether the name says what the figure is: Circle1 for "Circle", ParallelLine2 for "Parallel line"</summary>
        bool NameSays(string kind)
        {
            var name = Name.Replace(" ", "");
            return name.StartsWith(kind.Replace(" ", ""), StringComparison.OrdinalIgnoreCase)
                || name.StartsWith(GetType().Name, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// A figure whose name nobody reads (a measurement, a text, an angle's mark): with the
        /// number it is given, its <see cref="Title"/> is the kind alone, and the construction
        /// says which one it is ("Distance" "AB")
        /// </summary>
        protected virtual bool NamedByConstruction
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// How the figure is built, the words that follow its <see cref="Title"/>: "of CD" (a
        /// midpoint), "on circle k", "to line AB through E". Null where the title says it all
        /// (segment AB, a free point). The Figure List shows it after the title, faded, and
        /// the property grid under the title.
        /// </summary>
        public virtual string Construction
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// What another figure's <see cref="Construction"/> calls this one in front of its
        /// name: "segment", "line", "circle" - the plain word, not "parallel line". Null for a
        /// figure that goes by its name alone (a point, a number).
        /// </summary>
        public virtual string Noun
        {
            get
            {
                return Kind?.ToLowerInvariant();
            }
        }

        /// <summary>"circle c", "segment AB", "E": the figure as a construction names it (<see cref="ConstructionText.Of"/>)</summary>
        public string Reference
        {
            get
            {
                var name = NameDisplay.Format(Name);
                var noun = Noun;
                if (noun == null || string.IsNullOrEmpty(Name))
                {
                    return name;
                }

                return NameSays(noun) ? name : noun + " " + name;
            }
        }

        protected string mName;
        [PropertyGridVisible]
        [PropertyGridDisallowMultiEdit]
        [PropertyGridPreferredEditor("Name")]
        public virtual string Name
        {
            get
            {
                return mName;
            }
            set
            {
                // an empty name would not load: files refer to figures by name (the grid's
                // NameEditor refuses one; this is for everything else)
                if (string.IsNullOrEmpty(value))
                {
                    value = GenerateFigureName();
                }

                bool outermost = renameWave == null;
                if (outermost)
                {
                    renameWave = new Dictionary<IFigure, string>();
                }

                try
                {
                    if (!string.IsNullOrEmpty(mName) && mName != value && !renameWave.ContainsKey(this))
                    {
                        renameWave.Add(this, mName);
                    }

                    HasDefaultName = IsDefaultName(value);
                    mName = value;
                    if (Drawing != null && Drawing.Figures.Contains(this))
                    {
                        foreach (var f in Drawing.Figures.Where(f => f.Name == value).Where(f => f != this))
                        {
                            f.Name = f.GenerateFigureName(new List<string>() {this.Name});    // Rename figure with duplicate name.
                        }
                    }
                    RaisePropertyChanged("Name");
                    if (NameLabel != null && Drawing != null)
                    {
                        NameLabel.UpdateVisual();
                    }

                    foreach (var dependent in Dependents.OfType<FigureBase>().ToArray())
                    {
                        dependent.UpdateDefaultName();

                        // "of AB" names this one
                        dependent.RaiseConstructionChanged();
                    }
                }
                finally
                {
                    if (outermost)
                    {
                        var wave = renameWave;
                        renameWave = null;
                        if (!SuppressRenameInExpressions)
                        {
                            RenameInExpressions(wave);
                        }

                        SettleDefaultNames(wave);
                    }
                }
            }
        }

        /// <summary>
        /// The figures renamed by the rename under way, with the names they had: the one renamed,
        /// the figures named after it (segment AB follows point A), one that had to give up the
        /// new name. Null between renames.
        /// </summary>
        [ThreadStatic]
        static Dictionary<IFigure, string> renameWave;

        /// <summary>
        /// On while a point is put in place of another (<see cref="Actions.ReplacePoint"/>):
        /// the replacement passes through a temporary name, and the default names of what is
        /// built on it with it, before every name is back as it was - on undo too. Expressions
        /// wait that out; they would follow the temporary name and keep it.
        /// </summary>
        [ThreadStatic]
        public static bool SuppressRenameInExpressions;

        /// <summary>
        /// After a wave of renames, the expressions that name those figures say the new names
        /// (<see cref="ExpressionRenamer"/>), in everything built on them. Not recorded for undo,
        /// like the default names: undoing the rename renames back, and the text follows again.
        /// </summary>
        static void RenameInExpressions(Dictionary<IFigure, string> wave)
        {
            var renamed = wave
                .Where(pair => pair.Key.Name != pair.Value)
                .ToDictionary(pair => pair.Key, pair => pair.Value);
            var drawing = renamed.Keys.FirstOrDefault(figure => figure.Drawing != null)?.Drawing;
            if (drawing == null)
            {
                return;
            }

            var holders = DependencyAlgorithms
                .FindDescendants(f => f.Dependents, renamed.Keys)
                .OfType<IRenamableExpressions>()
                .ToArray();
            if (holders.Length == 0)
            {
                return;
            }

            var renamer = new ExpressionRenamer(drawing, renamed);
            foreach (var holder in holders)
            {
                holder.RenameInExpressions(renamer);
            }
        }

        public object Tag { get; set; }

        private bool mSelected = false;
        public virtual bool Selected
        {
            get
            {
                return mSelected;
            }
            set
            {
                mSelected = value;
            }
        }

        protected bool mEnabled = true;
        public virtual bool Enabled
        {
            get
            {
                return mEnabled;
            }
            set
            {
                mEnabled = value;
            }
        }

#if !PLAYER

        [PropertyGridVisible]
        [PropertyGridName("Style")]
        [PropertyGridGroup("Style")]
        [PropertyGridCustomValueProvider(typeof(StylePropertyValueProvider))]
        public virtual IFigureStyle StyleDisplay
        {
            get
            {
                return Style;
            }
            set
            {
                Style = value;
            }
        }

#endif

        protected bool mVisible = true;
        [PropertyGridVisible]
        public virtual bool Visible
        {
            get
            {
                return mVisible;
            }
            set
            {
                mVisible = value;
            }
        }

        public virtual bool IsHitTestVisible { get; set; }

        [PropertyGridVisible]
        public virtual bool Locked { get; set; }

        public bool Auxiliary { get; set; }

        public virtual void WriteXml(XmlWriter writer)
        {
            if (!Visible)
            {
                writer.WriteAttributeString("Visible", "false");
            }
            if (Locked)
            {
                writer.WriteAttributeString("Locked", "true");
            }
            if (Auxiliary)
            {
                writer.WriteAttributeBool("Auxiliary", true);
            }
            // the style of its kind goes without saying (a free point on FreePoint): a figure
            // without one gets it on loading (EnsureStyleAssigned)
            if (Style != null && Drawing != null && Style != Drawing.StyleManager.AssignDefaultStyle(this))
            {
                writer.WriteAttributeString("Style", Style.Name);
            }
            if (Flipped)
            {
                writer.WriteAttributeBool("Flipped", true);
            }
            if (!IsHitTestVisible)
            {
                writer.WriteAttributeBool("IsHitTestVisible", false);
            }
            if (Z != 0)
            {
                writer.WriteAttributeString("Z", Z.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        public virtual void ReadXml(XElement element)
        {
            Visible = element.ReadBool("Visible", true);
            Locked = element.ReadBool("Locked", false);
            Auxiliary = element.ReadBool("Auxiliary", false);
            IsHitTestVisible = element.ReadBool("IsHitTestVisible", true);
            Z = element.ReadInt("Z", 0);
            var styleAttribute = element.Attribute("Style");
            if (styleAttribute != null
                && Drawing != null
                && Drawing.StyleManager != null)
            {
                var style = Drawing.StyleManager[styleAttribute.Value];
                if (style != null)
                {
                    this.Style = style;
                }
            }
            Flipped = element.ReadBool("Flipped", false);
        }

        public virtual bool Serializable
        {
            get 
            { 
#if TABULA
                var mirror = Drawing.Figures.FirstOrDefault(f=>f is Mirror);
                if (mirror != null && this != mirror && this != (mirror as Mirror).Edge && 
                    (this.DependsOn(mirror) || this.DependsOn((mirror as Mirror).Edge)))
                {
                    return false;
                }
#endif
                return true; 
            }
        }

#if !PLAYER

        /// <summary>Over every other figure of its band of layers (<see cref="ZOrders"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Bring to front")]
        [PropertyGridIcon(PropertyGridIcon.BringToFront)]
        [PropertyGridCondition(nameof(CanBringToFront))]
        public void BringToFront()
        {
            ZOrders.BringToFront(new[] { (IFigure)this });
        }

        /// <summary>Under every other figure of its band of layers</summary>
        [PropertyGridVisible]
        [PropertyGridName("Send to back")]
        [PropertyGridIcon(PropertyGridIcon.SendToBack)]
        [PropertyGridCondition(nameof(CanSendToBack))]
        public void SendToBack()
        {
            ZOrders.SendToBack(new[] { (IFigure)this });
        }

        public bool CanBringToFront()
        {
            return ZOrders.CanBringToFront(new[] { (IFigure)this });
        }

        public bool CanSendToBack()
        {
            return ZOrders.CanSendToBack(new[] { (IFigure)this });
        }

        [PropertyGridVisible]
        [PropertyGridName("Delete")]
        [PropertyGridDestructive]
        public virtual void DeleteDisplay()
        {
            Actions.Remove(this);
        }

#endif

        protected Canvas Canvas
        {
            get
            {
                return Drawing?.Canvas;
            }
        }

        protected PointPair CanvasLogicalBorders
        {
            get
            {
                return ToLogical(Canvas.GetBorderRectangle());
            }
        }

        protected readonly ObservableCollection<IFigure> mDependencies = new ObservableCollection<IFigure>();
        public IList<IFigure> Dependencies
        {
            get
            {
                return mDependencies;
            }
            set
            {
                suppressDependencyListChangeNotification = true;
                try
                {
                    mDependencies.Clear();
                    if (value == null || value.Count == 0)
                    {
                        return;
                    }

                    mDependencies.AddRange(value);
                }
                finally
                {
                    suppressDependencyListChangeNotification = false;
                    OnDependenciesChanged();
                    UpdateDefaultName();
                }
            }
        }

        bool suppressDependencyListChangeNotification = false;

        private void mDependencies_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (!suppressDependencyListChangeNotification)
            {
                OnDependenciesChanged();
                UpdateDefaultName();
            }
        }

        protected virtual void OnDependenciesChanged()
        {
        }

        private IList<IFigure> mDependents;
        public IList<IFigure> Dependents
        {
            get
            {
                if (mDependents == null)
                {
                    mDependents = new List<IFigure>();
                }
                return mDependents;
            }
        }

        ZOrder layer;
        public ZOrder Layer
        {
            get
            {
                return layer;
            }
            set
            {
                layer = value;
                OnZIndexChanged();
            }
        }

        int z;
        public int Z
        {
            get
            {
                return z;
            }
            set
            {
                z = value;
                OnZIndexChanged();
            }
        }

        public int ZIndex
        {
            get
            {
                return ZOrders.Encode(Layer, Z);
            }
        }

        /// <summary>The layer or the Z changed: whatever is drawn for the figure takes the new ZIndex</summary>
        protected virtual void OnZIndexChanged()
        {
        }

        protected bool mExists = true;
        public virtual bool Exists
        {
            get
            {
                return mExists;
            }
            set
            {
                mExists = value;
            }
        }

        public virtual Point Point(int index)
        {
            if (index < mDependencies.Count && mDependencies[index] is IPoint point)
            {
                return point.Coordinates;
            }

            return this.MissingPoint(index);
        }

        IFigureStyle style;
        public virtual IFigureStyle Style
        {
            get
            {
                return style;
            }
            set
            {
                if (style == value)
                {
                    return;
                }
                if (style != null)
                {
                    style.PropertyChanged -= style_PropertyChanged;
                }
                style = value;
                if (style != null)
                {
                    style.PropertyChanged += style_PropertyChanged;
                }
                ApplyStyle();
            }
        }

#if !PLAYER

        [PropertyGridVisible]
        [PropertyGridName("Edit this style")]
        [PropertyGridGroup("Style")]
        [PropertyGridIcon(PropertyGridIcon.Pencil)]
        public void EditStyleButton()
        {
            var drawingHost = Canvas.Parent as DrawingHost;
            if (drawingHost != null && Style != null)
            {
                if (PropertyGrid != null)
                {
                    Style.CurrentEditInfo.ActionManager = this.Drawing.ActionManager;
                    Style.CurrentEditInfo.ParentObject = this;
                    Style.CurrentEditInfo.PropertyGrid = PropertyGrid;
                    PropertyGrid.Show(this.Style, this.Drawing.ActionManager);
                }
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Create new style")]
        [PropertyGridGroup("Style")]
        [PropertyGridIcon(PropertyGridIcon.Plus)]
        public void CreateNewStyle()
        {
            if (Style == null)
            {
                return;
            }

            // the new style and the figure taking it: one undo step
            using (Transaction.Create(Drawing.ActionManager, delayed: false))
            {
                Drawing.ActionManager.SetProperty(
                    this,
                    "Style",
                    Drawing.StyleManager.CreateNewStyle(this));
            }

            EditStyleButton();
        }

#endif

        void style_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // the drawing applies every style once at the end
            if (Drawing != null && Drawing.IsRefreshingTheme)
            {
                return;
            }

            ApplyStyle();
        }

        public void EnsureStyleAssigned()
        {
            if (Style == null && Drawing != null)
            {
                Style = Drawing.StyleManager.AssignDefaultStyle(this);
            }
        }

        public virtual void OnAddingToCanvas(Canvas newContainer)
        {
            EnsureStyleAssigned();
        }

        public abstract void ApplyStyle();

        public virtual void OnRemovingFromCanvas(Canvas leavingContainer)
        {
        }

        public virtual void UpdateExistence()
        {
            for (int i = 0; i < mDependencies.Count; i++)
            {
                if (!mDependencies[i].Exists)
                {
                    Exists = false;
                    return;
                }
            }

            Exists = true;
        }

        public virtual void Recalculate() { }

        /// <summary>
        /// Takes Coordinates or whatever other location information is current for the figure
        /// and updates the shape or other visual representation with these coordinates
        /// </summary>
        /// <example>
        /// Usually means updating the Shape like this:
        /// Shape.MoveTo(Coordinates.ToPhysical());
        /// </example>
        public virtual void UpdateVisual() { }

        public abstract IFigure HitTest(Point point);

        public bool Equals(IFigure other)
        {
            return object.ReferenceEquals(this, other);
        }

        public virtual Point Center
        {
            get
            {
                return new Point(0, 0);
            }
        }

        #region Coordinates

        protected double CursorTolerance
        {
            get
            {
                return Drawing.CoordinateSystem.CursorTolerance;
            }
        }

        protected double ToPhysical(double logicalLength)
        {
            return Drawing.CoordinateSystem.ToPhysical(logicalLength);
        }

        protected Point ToPhysical(Point point)
        {
            return Drawing.CoordinateSystem.ToPhysical(point);
        }

        protected PointPair ToPhysical(PointPair pointPair)
        {
            return Drawing.CoordinateSystem.ToPhysical(pointPair);
        }

        protected double ToLogical(double pixelLength)
        {
            return Drawing.CoordinateSystem.ToLogical(pixelLength);
        }

        protected Point ToLogical(Point pixel)
        {
            return Drawing.CoordinateSystem.ToLogical(pixel);
        }

        protected PointPair ToLogical(PointPair pointPair)
        {
            return Drawing.CoordinateSystem.ToLogical(pointPair);
        }

        #endregion

#if !PLAYER
        public PropertyGrid PropertyGrid { get; set; }
#endif

        public virtual object GetContentForPropertyGrid()
        {
            return this;
        }

        /// <summary>
        /// What <see cref="Construction"/> says has changed without a property of the figure's
        /// own (an expression of a point by coordinates edited, a point it is built on
        /// renamed): the grid's header reads it again
        /// </summary>
        public void RaiseConstructionChanged()
        {
            RaisePropertyChanged(nameof(Construction));
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public event PropertyChangedEventHandler PropertyChanging;
        protected void RaisePropertyChanging(string propertyName)
        {
            if (PropertyChanging != null)
            {
                PropertyChanging(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public bool Flipped { get; set; }
    }
}