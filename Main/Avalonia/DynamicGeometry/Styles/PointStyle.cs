using System.Collections.Generic;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    [StyleFor(typeof(IPoint))]
    public class PointStyle : ShapeStyle, IPropertyGridTabs
    {
        public override FrameworkElement GetSampleGlyph()
        {
            var point = Factory.CreatePointShape();
            point.Apply(this.GetWpfStyle());
            Resolve().OnApplied(null, point);

            // an emoji wants to be big on the canvas, but the row of styles is for telling them
            // apart, and every sample has a cell of the same size (StylePickerEditor.CellSize)
            double maxSize = Character != null ? MaxSampleCharacterSize : MaxSampleShapeSize;
            if (Size > maxSize)
            {
                point.Width = maxSize;
                point.Height = maxSize;
            }

            point.Tag = this;
            return point;
        }

        const double MaxSampleCharacterSize = 20;
        const double MaxSampleShapeSize = 24;

        PointShape shape = PointShape.Circle;
        [PropertyGridVisible]
        [PropertyGridPreferredEditor("PointShape")]
        public PointShape Shape
        {
            get
            {
                return shape;
            }
            set
            {
                shape = value;
                OnPropertyChanged("Shape");
            }
        }

        string character;

        /// <summary>
        /// Drawn instead of the shape (an emoji, or any one character): null for none, which is
        /// also what an empty text means. Declared before Size, which depends on it: a file is
        /// read in this order, and the Emoji tab shows the search first.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Emoji")]
        [PropertyGridPreferredEditor("Emoji")]
        public string Character
        {
            get
            {
                return character;
            }
            set
            {
                character = string.IsNullOrEmpty(value) ? null : value;
                OnPropertyChanged("Character");
                OnPropertyChanged("Size");
            }
        }

        /// <summary>A character is hard to make out at a shape's size</summary>
        public const double DefaultCharacterSize = 24;

        double shapeSize = 10.0;
        double characterSize = DefaultCharacterSize;

        /// <summary>
        /// Of the shape, or the height of the character, whichever the point shows. The other
        /// one is kept while the drawing is open, not saved: back on the other tab, the point
        /// gets its old size back (and undo needs nothing but the character).
        /// </summary>
        [PropertyGridVisible]
        [Domain(3, 300)]
        public double Size
        {
            get
            {
                return Character != null ? characterSize : shapeSize;
            }
            set
            {
                if (Character != null)
                {
                    characterSize = value;
                }
                else
                {
                    shapeSize = value;
                }

                OnPropertyChanged("Size");
            }
        }

        protected override void ApplyToWpfStyle(Style existingStyle, IFigure figure)
        {
            base.ApplyToWpfStyle(existingStyle, figure);
            existingStyle.Setters.Add(new Setter(FrameworkElement.WidthProperty, Size));
            existingStyle.Setters.Add(new Setter(FrameworkElement.HeightProperty, Size));
        }

        public override void OnApplied(IFigure figure, FrameworkElement element)
        {
            base.OnApplied(figure, element);
            if (element is PointMarker marker)
            {
                marker.Kind = Shape;
                marker.Character = Character;
            }
        }

        #region Tabs

        public const string ShapeTab = "Shape";
        public const string EmojiTab = "Emoji";

        static readonly string[] tabs = { ShapeTab, EmojiTab };

        IReadOnlyList<string> IPropertyGridTabs.Tabs => tabs;

        static readonly string[] shapeTab = { ShapeTab };
        static readonly string[] emojiTab = { EmojiTab };
        static readonly string[] belowTabs = { };

        IReadOnlyList<string> IPropertyGridTabs.GetTabs(string memberName)
        {
            switch (memberName)
            {
                case "Character":
                    return emojiTab;
                case "Size":
                    return tabs;
                case "FinishEditing":
                case "Delete":
                    return belowTabs;
                default:
                    return shapeTab;
            }
        }

        /// <summary>
        /// The character under the theme on screen, which is what the grid shows and edits
        /// (under any theme but the base one, the theme's own value). Going by the base
        /// value, an emoji picked under the dark theme opened on the Shape tab and could
        /// not be taken off again.
        /// </summary>
        string ShownCharacter
        {
            get
            {
                return ((PointStyle)Resolve()).Character;
            }
        }

        string IPropertyGridTabs.CurrentTab => ShownCharacter != null ? EmojiTab : ShapeTab;

        /// <summary>
        /// Back on Shape the character goes, so that the point is what the tab shows; on Emoji
        /// nothing changes until one is picked
        /// </summary>
        void IPropertyGridTabs.OnTabSelected(string tab, ActionManager actionManager)
        {
            if (tab != ShapeTab || ShownCharacter == null)
            {
                return;
            }

            var value = ThemedValue.ForCurrentTheme(PropertyDiscoveryStrategy.CreateValueProvider(this, "Character"));
            if (actionManager != null)
            {
                Actions.SetProperty(actionManager, value, "");
            }
            else
            {
                value.SetValue("");
            }
        }

        #endregion
    }
}
