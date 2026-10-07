using System;
using System.Linq;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    /// <summary>
    /// A check box on the paper that shows and hides figures: its dependencies, which it is
    /// not built on (it exists whatever they do) but holds. Made with the Show/hide box tool
    /// (<see cref="ShowHideCreator"/>), which also changes what one holds (Edit figures).
    /// </summary>
    public class ShowHideControl : ControlBase, ISupportRemoveDependency
    {
        public CheckBox Checkbox { get; set; }

        // the box on a plate that shows when it is selected, as a label's does (LabelBase)
        readonly Border plate = new Border();
        static readonly IBrush selectionBrush = new SolidColorBrush(Color.FromArgb(50, 51, 153, 255));

        // Fluent's check box template colors the caption by state through these resources,
        // over the control's own Foreground; and the outline of the empty box, which is the
        // theme's and was dark on a drawing's own dark paper (a night sky) under the light theme
        static readonly string[] captionResourceKeys =
        {
            "CheckBoxForegroundUnchecked",
            "CheckBoxForegroundUncheckedPointerOver",
            "CheckBoxForegroundUncheckedPressed",
            "CheckBoxForegroundChecked",
            "CheckBoxForegroundCheckedPointerOver",
            "CheckBoxForegroundCheckedPressed",
            "CheckBoxForegroundIndeterminate",
            "CheckBoxForegroundIndeterminatePointerOver",
            "CheckBoxForegroundIndeterminatePressed",
            "CheckBoxCheckBackgroundStrokeUnchecked",
            "CheckBoxCheckBackgroundStrokeUncheckedPointerOver",
            "CheckBoxCheckBackgroundStrokeUncheckedPressed"
        };

        protected override string Kind
        {
            get
            {
                return "Show/hide box";
            }
        }

        /// <summary>The caption, in quotes</summary>
        public override string Construction
        {
            get
            {
                return ConstructionText.Quote(Checkbox?.Content?.ToString());
            }
        }

        /// <summary>The caption and the empty box in the text style's color, in every state (hovered, pressed, checked)</summary>
        public override void ApplyStyle()
        {
            if (Style == null)
            {
                return;
            }

            this.Apply(Checkbox, Style);
            foreach (var key in captionResourceKeys)
            {
                Checkbox.Resources[key] = Checkbox.Foreground;
            }

            plate.Background = Selected ? selectionBrush : null;
            base.ApplyStyle();
        }

        /// <summary>The caption beside the box</summary>
        [PropertyGridVisible]
        [PropertyGridName("Caption")]
        public string Text
        {
            get
            {
                return Checkbox.Content?.ToString() ?? "";
            }
            set
            {
                Checkbox.Content = value ?? "";

                // a pinned box keeps its corner as the caption changes its size
                if (Drawing != null)
                {
                    UpdateVisual();
                }

                // (the grid's title quotes the caption)
                RaisePropertyChanged(nameof(Text));
            }
        }

        /// <summary>What the box shows and hides, in words</summary>
        [PropertyGridVisible]
        [PropertyGridName("Shows and hides")]
        public string ShownFigures
        {
            get
            {
                return Describe(Dependencies);
            }
        }

        /// <summary>"segment AB, point C and 3 more"; "nothing" for none</summary>
        public static string Describe(System.Collections.Generic.IList<IFigure> figures)
        {
            const int shown = 4;
            if (figures.Count == 0)
            {
                return "nothing";
            }

            var names = string.Join(", ", figures.Take(shown).Select(ConstructionText.Of));
            return figures.Count > shown ? names + " and " + (figures.Count - shown) + " more" : names;
        }

        /// <summary>Picks again what the box shows and hides, with the Show/hide box tool</summary>
        [PropertyGridVisible]
        [PropertyGridName("Edit figures")]
        [PropertyGridIcon(PropertyGridIcon.Pencil)]
        public void EditFigures()
        {
            ShowHideCreator.Edit(this);
        }

        /// <summary>
        /// A box is dragged by itself: what it holds is not what it is built on (the base
        /// would move the points of those figures instead, as for a figure built on points)
        /// </summary>
        public override bool AllowMove()
        {
            return !Locked;
        }

        /// <summary>
        /// A figure the box holds, deleted, leaves the box with the others; the last one takes
        /// the box along. (A box that held several figures built on the point deleted asks
        /// for each before any has gone, and can be left holding nothing.)
        /// </summary>
        public bool CanRemoveDependency(IFigure dependency)
        {
            return Dependencies.Count > 1 && Dependencies.Contains(dependency);
        }

        public IAction GetRemoveDependencyAction(IFigure dependency)
        {
            int index = Dependencies.IndexOf(dependency);
            return new CallMethodAction(
                () => this.RemoveDependencyCore(index, dependency),
                () => this.InsertDependencyCore(index, dependency));
        }

        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);

            // Only the box: each figure says in the file whether it is hidden. Hiding them
            // all again here undid what the user had done since the box was last clicked -
            // one of its figures shown by hand came back hidden when the file was opened.
            SetBox(element.ReadBool("Show", true));
            Checkbox.Content = element.ReadString("Text");
            var pinName = element.ReadString("Pin");
            if (pinName != null && Enum.TryParse(pinName, out LabelPin readPin) && readPin != LabelPin.None)
            {
                PinOffset = new Point(element.ReadDouble("OffsetX"), element.ReadDouble("OffsetY"));
                Pin = readPin;
                UpdateVisual();
            }
            else
            {
                var x = element.ReadDouble("X");
                var y = element.ReadDouble("Y");
                MoveTo(new Point(x, y));
            }
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeBool("Show", Checkbox.IsChecked == true);
            writer.WriteAttributeString("Text", Text);
            if (Pin == LabelPin.None)
            {
                var coordinates = Coordinates;
                writer.WriteAttributeString("X", coordinates.X.ToStringInvariant());
                writer.WriteAttributeString("Y", coordinates.Y.ToStringInvariant());
            }
            else
            {
                writer.WriteAttributeString("Pin", Pin.ToString());
                writer.WriteAttributeDouble("OffsetX", PinOffset.X);
                writer.WriteAttributeDouble("OffsetY", PinOffset.Y);
            }
        }

        #region Pinning

        /// <summary>
        /// A box pinned to a corner of the canvas stays there in pixels while the plane zooms
        /// and pans, as a pinned label does (<see cref="Label.Pin"/>). Boxes stacked at places
        /// in the plane were a different number of pixels apart at every zoom: overlapping on
        /// a phone, far apart on a big screen. Only a file pins a box.
        /// </summary>
        public LabelPin Pin { get; set; }

        /// <summary>Pixels from the pinned corner of the canvas to the same corner of the box</summary>
        public Point PinOffset { get; set; }

        bool HasCanvas
        {
            get
            {
                return Drawing != null && Drawing.Canvas != null;
            }
        }

        /// <summary>
        /// The size of the box in pixels, measured now: right after a load there was no layout
        /// pass yet. The plate's measure is invalidated first: it stays valid when only the
        /// caption inside it changed.
        /// </summary>
        public Size MeasureSize()
        {
            Shape.InvalidateMeasure();
            Shape.Measure(Size.Infinity);
            return Shape.DesiredSize;
        }

        /// <summary>The top-left corner of a pinned box on the canvas, in pixels</summary>
        public Point PinnedTopLeft()
        {
            return Pinning.TopLeft(Pin, PinOffset, MeasureSize(), Drawing.CoordinateSystem.PhysicalSize);
        }

        public override void UpdateVisual()
        {
            if (Pin == LabelPin.None)
            {
                base.UpdateVisual();
                return;
            }

            if (!HasCanvas)
            {
                return;
            }

            var topLeft = PinnedTopLeft();
            Coordinates = ToLogical(topLeft);
            Shape.MoveTo(topLeft);
        }

        /// <summary>Dragging a pinned box changes its offset from the corner, not its place in the plane</summary>
        public override void MoveToCore(Point newLocation)
        {
            if (Pin != LabelPin.None && HasCanvas)
            {
                PinOffset = Pinning.OffsetFrom(Pin, ToPhysical(newLocation), MeasureSize(), Drawing.CoordinateSystem.PhysicalSize);
            }

            base.MoveToCore(newLocation);
        }

        /// <summary>A pinned box's place is its offset from the corner, in pixels; an unpinned one's is in the plane</summary>
        public override object CapturePlace()
        {
            return Pin != LabelPin.None ? new PinnedPlace() { Offset = PinOffset } : base.CapturePlace();
        }

        public override void RestorePlace(object place)
        {
            if (place is PinnedPlace pinned)
            {
                PinOffset = pinned.Offset;
                if (HasCanvas)
                {
                    UpdateVisual();
                }
            }
            else
            {
                base.RestorePlace(place);
            }
        }

        class PinnedPlace
        {
            public Point Offset;
        }

        #endregion

        protected override FrameworkElement CreateShape()
        {
            Checkbox = new ShowHideCheckBox() { LeavesPressToTool = () => Drawing?.Behavior is Dragger or ShowHideCreator };

            // no plate of its own on the canvas; the theme paints the hover
            Checkbox.Background = Brushes.Transparent;
            Checkbox.IsCheckedChanged += (s, e) =>
            {
                if (!settingBox)
                {
                    Toggle(Checkbox.IsChecked == true);
                }
            };
            plate.Child = Checkbox;
            return plate;
        }

        /// <summary>
        /// The box on the canvas. Under the Drag tool a press on it is the tool's (it bubbles
        /// on to the canvas): a click ticks the box there (<see cref="Click"/>) and a drag
        /// moves it, as a drag moves any figure. Taken by the check box itself, as under the
        /// other tools, the press never reached the tool, and nothing could move a box. Under
        /// the Show/hide box tool too, where a click on a box ends the editing of its figures
        /// and ticks nothing.
        /// </summary>
        public class ShowHideCheckBox : CheckBox
        {
            public Func<bool> LeavesPressToTool { get; set; }

            protected override Type StyleKeyOverride => typeof(CheckBox);

            protected override void OnPointerPressed(PointerPressedEventArgs e)
            {
                if (LeavesPressToTool?.Invoke() == true)
                {
                    return;
                }

                base.OnPointerPressed(e);
            }

            protected override void OnPointerReleased(PointerReleasedEventArgs e)
            {
                if (LeavesPressToTool?.Invoke() == true)
                {
                    return;
                }

                base.OnPointerReleased(e);
            }
        }

        /// <summary>A click on the box that reached the Drag tool: ticks or unticks it, an undo step</summary>
        public void Click()
        {
            bool show = Checkbox.IsChecked != true;
            SetBox(show);
            Toggle(show);
        }

        // the box is being ticked by the program (undo, redo): not a click to record
        bool settingBox;

        /// <summary>
        /// A click on the box. The box and what it shows are saved with the drawing, so this
        /// is an undo step: undo puts the box back and gives each figure the visibility it
        /// had, which need not have been what the box said.
        /// </summary>
        void Toggle(bool show)
        {
            // from a file, or in the middle of a construction, whose undo step is its figure's alone
            if (Drawing == null || !Drawing.Figures.Contains(this) || Drawing.IsRecordingTransaction)
            {
                Show(show);
                return;
            }

            var figures = Dependencies.ToArray();
            var before = figures.Select(figure => figure.Visible).ToArray();
            Drawing.ActionManager.RecordAction(new CallMethodAction(
                () =>
                {
                    SetBox(show);
                    Show(show);
                },
                () =>
                {
                    SetBox(!show);
                    for (int i = 0; i < figures.Length; i++)
                    {
                        figures[i].Visible = before[i];
                        figures[i].UpdateVisual();
                    }
                }));
        }

        /// <summary>Ticks or unticks the box, leaving its figures as they are</summary>
        public void SetBox(bool isChecked)
        {
            settingBox = true;
            try
            {
                Checkbox.IsChecked = isChecked;
            }
            finally
            {
                settingBox = false;
            }
        }

        public void UpdateFigureVisibility()
        {
            Show(Checkbox.IsChecked == true);
        }

        private void Show(bool show)
        {
            foreach (var figure in Dependencies)
            {
                figure.Visible = show;
                figure.UpdateVisual();
            }
        }

        /// <summary>
        /// The box stays whether or not its figures exist right now: they are what it shows and
        /// hides, not what it is built on. Taken for dependencies, a hint that was nowhere at the
        /// moment (a firework that hadn't burst yet) took the box along.
        /// </summary>
        public override void UpdateExistence()
        {
            Exists = true;
        }
    }
}
