using System.Collections;
using System.Windows.Media;

namespace DynamicGeometry
{
    public partial class Settings
    {
        // Developer Settings
        public static bool ChangeLineAppearanceWhenSelected = true;
        public static bool ChangePointAppearanceWhenSelected = true;
        public static bool ChangePointStrokeWidthWhenSelected = false;
        public static double DefaultUnitLength = 48;
        public static double DefaultToolbarFontSize = 11;
        public static Color PropertyGridTitleColor = Colors.Green;
        public static bool ShowIconInTabPanelHeader = true;
        public static bool UpdateSelectedBehaviorOnTabChange = true;
        public static bool ShowStyleNameInStylePicker = false;
        public static bool ScaleTextWithDrawing = false;
        public static double CurrentDrawingVersion = 0;

        /// <summary>
        /// Should the orientation of a PointOnFigure on an elliptical figure remain fixed or should it be relative to the orientation of the figure?
        /// </summary>
        public static bool PointsOnEllipticalsUseAbsoluteAngle = true;

        static Settings instance = new Settings();
        public static Settings Instance
        {
            get
            {
                return instance;
            }
            set
            {
                instance = value;
            }
        }

        public virtual bool AutoLabelPoints { get; set; }
        public virtual bool ShowGrid { get; set; }
        public virtual bool ShowFigureExplorer { get; set; }
        public virtual bool EnableOrtho { get; set; }
        public virtual bool EnableSnapToGrid { get; set; }
        public virtual bool EnableSnapToPoint { get; set; }
        public virtual bool EnableSnapToCenter { get; set; }
        public virtual bool HideHints { get; set; }
        public virtual Math.lengthUnit DistanceUnit { get; set; } // Used by Measurement subclasses.  Not yet implemented throughout.

        private double snapGridSpacing = 1;
        public virtual double SnapGridSpacing
        {
            get 
            { 
                return snapGridSpacing; 
            }
            set 
            { 
                snapGridSpacing = value; 
            }
        }

        private double cursorTolerance = 5;
        public virtual double CursorTolerance
        {
            get
            {
                return cursorTolerance;
            }
            set
            {
                cursorTolerance = value;
            }
        }

        string pointAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public virtual string PointAlphabet
        {
            get
            {
                return pointAlphabet;
            }
            set
            {
                pointAlphabet = value;
            }
        }

        // Polar
        public virtual bool EnablePolar { get; set; }
        public virtual PolarValue PolarIncrement { get; set; }  // current polar increment
        public virtual string TempNewIncrement { get; set; }    // polar increment to add in IEnumerable PolarItems
        public virtual IEnumerable PolarItems { get; set; }

        // Length and Angle which were input by User in Traction Mode
        public virtual bool UserModeLength { get; set; }
        public virtual bool UserModeAngle { get; set; }
        public virtual double UserAngle { get; set; }
        public virtual double UserLength { get; set; }
    }
}