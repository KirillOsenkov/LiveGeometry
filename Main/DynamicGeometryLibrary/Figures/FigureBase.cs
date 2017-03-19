using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
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
            return this.GenerateNewName();
        }

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

        protected string mName;
        [PropertyGridVisible]
        [PropertyGridDisallowMultiEdit]
        public virtual string Name
        {
            get
            {
                return mName;
            }
            set
            {
                mName = value;
                if (Drawing != null && Drawing.Figures.Contains(this))
                {
                    foreach (var f in Drawing.Figures.Where(f => f.Name == value).Where(f => f != this))
                    {
                        f.Name = f.GenerateFigureName(new List<string>() {this.Name});    // Rename figure with duplicate name.
                    }
                }
                RaisePropertyChanged("Name");
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
        [PropertyGridCustomValueProvider(typeof(StylePropertyValueProvider))]
        public IFigureStyle StyleDisplay
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
            if (Style != null)
            {
                writer.WriteAttributeString("Style", Style.Name);
            }
            if (Flipped)
            {
                writer.WriteAttributeBool("Flipped", true);
            }
        }

        public virtual void ReadXml(XElement element)
        {
            Visible = element.ReadBool("Visible", true);
            Locked = element.ReadBool("Locked", false);
            IsHitTestVisible = element.ReadBool("IsHitTestVisible", true);
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

        [PropertyGridVisible]
        [PropertyGridName("Delete")]
        public virtual void DeleteDisplay()
        {
            Actions.Remove(this);
        }

#endif

        protected Canvas Canvas
        {
            get
            {
                return Drawing.Canvas;
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
                }
            }
        }

        bool suppressDependencyListChangeNotification = false;

        private void mDependencies_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (!suppressDependencyListChangeNotification)
            {
                OnDependenciesChanged();
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

        public virtual int ZIndex { get; set; }

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
            var point = mDependencies[index] as IPoint;
            return point.Coordinates;
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
        [PropertyGridName("Edit style")]
        public void EditStyleButton()
        {
            var drawingHost = Canvas.Parent as DrawingHost;
            if (drawingHost != null)
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
        public void CreateNewStyle()
        {
            Drawing.ActionManager.SetProperty(
                this,
                "Style",
                Drawing.StyleManager.CreateNewStyle(this));
            EditStyleButton();
        }

#endif

        void style_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
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
            Canvas.SetZIndex(newContainer, this.ZIndex);
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