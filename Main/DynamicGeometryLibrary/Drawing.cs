using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public partial class Drawing
    {

        public Drawing(Canvas canvas)
        {
            Check.NotNull(canvas, "canvas");

            ActionManager = new ActionManager();
            StyleManager = new StyleManager(this);

            Figures = new RootFigureList(this);

            OnAttachToCanvas += Drawing_OnAttachToCanvas;
            OnDetachFromCanvas += Drawing_OnDetachFromCanvas;

            Canvas = canvas;

            CoordinateSystem = new CoordinateSystem(this);
            CoordinateGrid = new CartesianGrid() { Drawing = this, Visible = Settings.Instance.ShowGrid };
            Figures.Add(CoordinateGrid);
            Version = Settings.CurrentDrawingVersion;
        }

        public double Version { get; set; }

        void Drawing_OnAttachToCanvas(Canvas canvas)
        {
            canvas.SizeChanged += mCanvas_SizeChanged;
            UpdateClip(canvas);
            foreach (var figure in Figures)
            {
                figure.OnAddingToCanvas(canvas);
            }
        }

        void UpdateClip(Canvas canvas)
        {
            canvas.Clip = new RectangleGeometry() { Rect = new Rect(0, 0, canvas.ActualWidth, canvas.ActualHeight) };
        }

        void Drawing_OnDetachFromCanvas(Canvas canvas)
        {
            canvas.SizeChanged -= mCanvas_SizeChanged;
            foreach (var figure in Figures)
            {
                figure.OnRemovingFromCanvas(canvas);
            }
            this.Behavior = null;
        }

        #region Events

        public event EventHandler<SelectionChangedEventArgs> SelectionChanged;

        public class SelectionChangedEventArgs : EventArgs
        {
            public SelectionChangedEventArgs()
            {
                SelectedFigures = Enumerable.Empty<IFigure>();
            }

            public SelectionChangedEventArgs(IEnumerable<IFigure> selection)
                : this()
            {
                SelectedFigures = selection;
            }

            public SelectionChangedEventArgs(IFigure singleSelection)
                : this(singleSelection.AsEnumerable())
            {
            }

            public IEnumerable<IFigure> SelectedFigures { get; set; }
        }

        public event EventHandler<DeleteExecutedEventArgs> DeleteExecuted;

        public class DeleteExecutedEventArgs : EventArgs
        {
            public DeleteExecutedEventArgs()
            {
                DeletedFigures = Enumerable.Empty<IFigure>();
            }

            public DeleteExecutedEventArgs(IEnumerable<IFigure> deletedFigures)
                : this()
            {
                DeletedFigures = deletedFigures;
            }

            public DeleteExecutedEventArgs(IFigure deletedFigure)
                : this(deletedFigure.AsEnumerable())
            {
            }

            public IEnumerable<IFigure> DeletedFigures { get; set; }
        }

        public void ClearLockedFigures()
        {
            foreach (IFigure figure in this.Figures)
            {
                if (figure.Locked)
                {
                    figure.Locked = false;
                }
            }
        }

        public void RaiseSelectionChanged(SelectionChangedEventArgs args)
        {
            if (SelectionChanged != null)
            {
                SelectionChanged(this, args);
            }
        }

        public void RaiseSelectionChanged(params IFigure[] selected)
        {
            if (SelectionChanged != null)
            {
                SelectionChanged(this, new SelectionChangedEventArgs(selected));
            }
        }

        public void RaiseSelectionChanged(IEnumerable<IFigure> selected)
        {
            if (SelectionChanged != null)
            {
                SelectionChanged(this, new SelectionChangedEventArgs(selected));
            }
        }

        public void RaiseDeleteExecuted(DeleteExecutedEventArgs args)
        {
            if (DeleteExecuted != null)
            {
                DeleteExecuted(this, args);
            }
        }

        public void RaiseDeleteExecuted(params IFigure[] selected)
        {
            if (DeleteExecuted != null)
            {
                DeleteExecuted(this, new DeleteExecutedEventArgs(selected));
            }
        }

        public void RaiseDeleteExecuted(IEnumerable<IFigure> selected)
        {
            if (DeleteExecuted != null)
            {
                DeleteExecuted(this, new DeleteExecutedEventArgs(selected));
            }
        }

        public class DisplayPropertiesEventArgs : EventArgs
        {
            public object Object { get; set; }
        }

        public event EventHandler<DisplayPropertiesEventArgs> DisplayProperties;

        public void RaiseDisplayProperties(object objectWithProperties)
        {
            if (DisplayProperties != null)
            {
                DisplayProperties(this, new DisplayPropertiesEventArgs() { Object = objectWithProperties });
            }
        }

        public class ConstructionStepCompleteEventArgs : EventArgs
        {
            public bool ConstructionComplete { get; set; }
            public bool ConstructionRollback { get; set; }
            public Type FigureTypeNeeded { get; set; }
            public IEnumerable<IFigure> FigureResults { get; set; }
        }

        public class ConstructionStepStartedEventArgs : EventArgs
        {
        }

        public event EventHandler<ConstructionStepStartedEventArgs> ConstructionStepStarted;
        public event EventHandler<ConstructionStepCompleteEventArgs> ConstructionStepComplete;

        public void RaiseConstructionStepComplete(ConstructionStepCompleteEventArgs args)
        {
            if (ConstructionStepComplete != null)
            {
                ConstructionStepComplete(this, args);
            }
        }

        public void RaiseConstructionStepStarted(ConstructionStepStartedEventArgs args)
        {
            if (ConstructionStepStarted != null)
            {
                ConstructionStepStarted(this, args);
            }
        }

        public void RaiseConstructionStepStarted()
        {
            if (ConstructionStepStarted != null)
            {
                ConstructionStepStarted(this, new ConstructionStepStartedEventArgs());
            }
        }

        public class ConstructionFeedbackEventArgs : EventArgs
        {
            public Type FigureTypeNeeded { get; set; }
            public bool IsMouseButtonDown { get; set; }
        }

        public event EventHandler<ConstructionFeedbackEventArgs> ConstructionFeedback;
        public void RaiseConstructionFeedback(ConstructionFeedbackEventArgs args)
        {
            if (ConstructionFeedback != null)
            {
                ConstructionFeedback(this, args);
            }
        }

        // UserIsAddingFigures is intended to occur after figures have been added to drawing but before transaction is committed.
        public event EventHandler<UIAFEventArgs> UserIsAddingFigures;

        public class UIAFEventArgs : EventArgs
        {
            public IEnumerable<IFigure> Figures { get; set; }
        }

        /// <summary>
        /// This offers an opportunity to do additional processing of figures after they have been added to the drawing.
        /// Unlike FigureList.OnItemAdded(), this event does not occur on undo or redo.
        /// </summary>
        public void RaiseUserIsAddingFigures(UIAFEventArgs figures)
        {
            if (UserIsAddingFigures != null)
            {
                UserIsAddingFigures(this, figures);
            }
        }

        public class DocumentOpenRequestedEventArgs : EventArgs
        {
            public enum InWhichWindowChoice
            {
                DontCare,
                ReuseCurrent,
                NewWindowOrTab
            }

            public string DocumentXml { get; set; }
            public InWhichWindowChoice InWhichWindow { get; set; }
        }

        /// <summary>
        /// This event is raised when the user clicks on a hyperlink
        /// to open another drawing document, much like a web-browser link.
        /// This event signals to the host of the drawing to either open this 
        /// new drawing in a separate tab or replace the current one.
        /// </summary>
        public event EventHandler<DocumentOpenRequestedEventArgs> DocumentOpenRequested;
        public void RaiseDocumentOpenRequested(DocumentOpenRequestedEventArgs args)
        {
            if (DocumentOpenRequested != null)
            {
                DocumentOpenRequested(this, args);
            }
        }

        public event SizeChangedEventHandler SizeChanged;
        public void RaiseSizeChanged(SizeChangedEventArgs args)
        {
            if (SizeChanged != null)
            {
                SizeChanged(this, args);
            }
        }

        public event Action<string> Status;
        public void RaiseStatusNotification(string status)
        {
            if (Status != null)
            {
                Status(status);
            }
        }

        public event Action ZoomChanged;    // Used by Tabula.
        public void RaiseZoomChanged()
        {
            if (ZoomChanged != null)
            {
                ZoomChanged();
            }
        }

        public event EventHandler<FigureCoordinatesChangedEventArgs> FigureCoordinatesChanged;
        public class FigureCoordinatesChangedEventArgs : EventArgs
        {
            public FigureCoordinatesChangedEventArgs()
            {
                Figures = Enumerable.Empty<IFigure>();
            }

            public FigureCoordinatesChangedEventArgs(IEnumerable<IFigure> figures)
                : this()
            {
                Figures = figures;
            }

            public FigureCoordinatesChangedEventArgs(IFigure singleFigure)
                : this(singleFigure.AsEnumerable())
            {
            }

            public IEnumerable<IFigure> Figures { get; set; }
        }

        public void RaiseFigureCoordinatesChanged(FigureCoordinatesChangedEventArgs args)
        {
            if (FigureCoordinatesChanged != null)
            {
                FigureCoordinatesChanged(this, args);
            }
        }

        #endregion

        public string Name { get; set; }
        public override string ToString()
        {
            return Name;
        }

        /// <summary>
        /// Compares the last item in ActionManager.EnumUndoableActions() to a value stored at the last Save to determine if there are unsaved changes.
        /// </summary>
        public bool HasUnsavedChanges
        {
            get
            {
                var undoableActions = ActionManager.EnumUndoableActions();
                if (undoableActions.IsEmpty())
                {
                    return false;
                }

                if (LastUndoableActionAtSave == undoableActions.Last())
                {
                    return false;
                }
                else
                {
                    return true;
                }

            }
        }

        public IAction LastUndoableActionAtSave { get; set; }
        public ActionManager ActionManager { get; set; }

        public event Action<Canvas> OnAttachToCanvas;
        public event Action<Canvas> OnDetachFromCanvas;

        public event Action<Behavior> BehaviorChanged;

        /// <summary>
        /// Informs the clients that the current behavior of the drawing
        /// was set to a new behavior.
        /// </summary>
        /// <param name="behavior">The new behavior of the drawing.</param>
        void RaiseBehaviorChanged(Behavior behavior)
        {
            if (BehaviorChanged != null)
            {
                BehaviorChanged(behavior);
            }
        }

        private Canvas mCanvas;
        public Canvas Canvas
        {
            get
            {
                return mCanvas;
            }
            set
            {
                if (mCanvas == value)
                {
                    return;
                }
                if (mCanvas != null && OnDetachFromCanvas != null)
                {
                    OnDetachFromCanvas(mCanvas);
                }
                mCanvas = value;
                if (mCanvas != null && OnAttachToCanvas != null)
                {
                    OnAttachToCanvas(mCanvas);
                }
            }
        }

        void mCanvas_SizeChanged(object sender, System.Windows.SizeChangedEventArgs e)
        {
            UpdateClip(Canvas);
            RaiseSizeChanged(e);
        }

        public StyleManager StyleManager { get; set; }

        private Behavior mBehavior;
        public Behavior Behavior
        {
            get
            {
                return mBehavior;
            }
            set
            {
                if (mBehavior == value)
                {
                    return;
                }
                if (mBehavior != null)
                {
                    mBehavior.Stopping();
                    mBehavior.Drawing = null;
                }
                mBehavior = value;
                if (mBehavior != null)
                {
                    mBehavior.Drawing = this;
                    mBehavior.Started();
                    RaiseBehaviorChanged(mBehavior);
                }
                
            }
        }

        public void SetDefaultBehavior()
        {
            Behavior = Behavior.Default;
        }

        public CoordinateSystem CoordinateSystem { get; set; }
        public CartesianGrid CoordinateGrid { get; set; }

        public RootFigureList Figures { get; set; }

        public IEnumerable<IFigure> GetSerializableFigures()
        {
            foreach (IFigure figure in Figures.Where(f=>f.Serializable))
            {
                yield return figure;
            }
        }

        public IEnumerable<IFigure> GetSelectedFigures()
        {
            foreach (IFigure figure in Figures)
            {
                if (figure.Selected)
                {
                    yield return figure;
                }
            }
        }

        public Point GetSelectionCenter()
        {
            Point center = new Point(0, 0);
            var selectedFigures = GetSelectedFigures();
            var count = selectedFigures.Count();
            if (count > 0)
            {
                foreach (var figure in selectedFigures)
                {
                    center.X += figure.Center.X;
                    center.Y += figure.Center.Y;
                }
                center.X /= count;
                center.Y /= count;
            }
            return center;
        }

        public List<IFigure> GetSelectedFiguresWithDependencies()
        {
            var selectedFigures = GetSelectedFigures();
            List<IFigure> results = new List<IFigure>();
            results.AddRange(selectedFigures);
            foreach (IFigure selectedFigure in selectedFigures)
            {
                foreach (IFigure f in Figures)
                {
                    if (selectedFigure.DependsOn(f) && !results.Contains(f))
                    {
                        results.Add(f);
                    }
                }
            }
            return results;
        }

        public IEnumerable<IFigure> GetLockedFigures()
        {
            foreach (IFigure figure in Figures)
            {
                if (figure.Locked)
                {
                    yield return figure;
                }
            }
        }

#if !SILVERLIGHT
        public static Drawing Load(string path, Canvas canvas)
        {
            Drawing drawing = new Drawing(canvas);
            new DrawingDeserializer().ReadDrawing(drawing, System.IO.File.ReadAllText(path));
            return drawing;
        }

        public void Save(string path)
        {
            DrawingSerializer.Save(this, path);
        }
#endif

        [Obsolete("Use Actions.Add instead")]
        public void Add(IFigure figure)
        {
            Actions.Add(this, figure);
        }

#if !PLAYER

        [Obsolete("Use Actions.Remove instead")]
        public void Remove(IFigure figure)
        {
            Actions.Remove(figure);
        }

#endif

        public void Recalculate()
        {
            foreach (var figure in Figures)
            {
                figure.RecalculateAndUpdateVisual();
            }
        }

        [Obsolete("Use Actions.Add instead")]
        public void Add(IEnumerable<IFigure> figures)
        {
            using (Transaction.Create(ActionManager))
            {
                foreach (var figure in figures)
                {
                    Actions.Add(this, figure);
                }
            }
        }

        public void AddFromXml(XElement element)
        {
            var deserializer = new DrawingDeserializer();
            deserializer.ReadDrawing(this, element);
            if (!deserializer.IsSuccess)
            {
                RaiseStatusNotification(deserializer.GetErrorReport());
            }
        }

#if !TABULAPLAYER
        public void AddFromDGF(string[] lines)
        {
            var reader = new DGFReader();
            reader.ReadDrawing(this, lines);
            if (!reader.IsSuccess)
            {
                RaiseStatusNotification(reader.GetErrorReport());
            }
        }
#endif
#if !PLAYER

        public string SaveAsText()
        {
            return DrawingSerializer.SaveDrawing(this);
        }

        public void DeleteSelection()
        {
            Actions.RemoveMany(this, this.GetSelectedFigures().Where(f => !(f is CartesianGrid)).TopologicalSort(f => f.Dependents).Where(f => !(f is PointLabel)));
        }

#endif

#if !SILVERLIGHT

        public void Copy()
        {
            List<IFigure> list = new List<IFigure>(this.GetSelectedFiguresWithDependencies());
            var s = new System.Text.StringBuilder();
            using (var w = System.Xml.XmlWriter.Create(s, new System.Xml.XmlWriterSettings()
            {
                Indent = true
            }))
            {
                new DrawingSerializer().WriteFigureList(list, w);
            }
            Clipboard.SetText(s.ToString()); 
        }

        public void Paste()
        {
            if (Clipboard.GetText() != null)
            {
                this.PasteFromText(Clipboard.GetText());
            }
        }

        public void PasteFromText(string str)
        {  
            if (str != null)
            {
                Actions.Paste(this, str);
            }
        }

        public void PasteFrom(string xmlFile)
        {
            string copiedFigures = System.IO.File.ReadAllText(xmlFile);
            PasteFromText(copiedFigures);
        }

#endif
#if !PLAYER
        public void Duplicate()
        {
            var transaction = Transaction.Create(ActionManager, false);
            List<IFigure> list = GetSelectedFiguresWithDependencies();
            var s = new System.Text.StringBuilder();
            using (var w = System.Xml.XmlWriter.Create(s, new System.Xml.XmlWriterSettings()
            {
                Indent = true
            }))
            {
                new DrawingSerializer().WriteFigureList(list, w);
            }
            var paste = new PasteAction(this, s.ToString());
            ActionManager.RecordAction(paste);

            // Offset the new figures.
            List<IMovable> moving = new List<IMovable>();
            foreach (IFigure f in paste.Figures.Where(f => f as IMovable != null))
            {
                moving.Add(f as IMovable);
            }
            Actions.Move(this, moving, new Point(1, -1), Figures);

            // Select new figures.
            Figures.ClearSelection();
            foreach (IFigure f in paste.Figures)
            {
                f.Selected = true;
            }

            RaiseUserIsAddingFigures(new UIAFEventArgs() {Figures = paste.Figures});
            transaction.Commit();
        }

#endif
      
        public void SelectAll()
        {
            foreach (IFigure figure in Figures)
            {
                if ((!(figure is CartesianGrid)))
                {
                    figure.Selected = true;
                }
            }
        }

        public void LockSelected()
        {
            IEnumerable<IFigure> roots = this.GetSelectedFigures();
            bool shouldLock = roots.All(root => (!root.Locked));
            foreach (IFigure figure in roots)
            {
                if (!(figure is CartesianGrid))
                {
                    figure.Locked = shouldLock;
                }
            }
        }

        public void ClearStatus()
        {
            RaiseStatusNotification("");
        }

        public event EventHandler<UnhandledExceptionNotificationEventArgs> UnhandledException;
        public void RaiseError(object sender, Exception ex)
        {
            if (UnhandledException != null)
            {
                UnhandledException(sender, new UnhandledExceptionNotificationEventArgs(ex));
            }
        }
    }

    public class UnhandledExceptionNotificationEventArgs : EventArgs
    {
        public UnhandledExceptionNotificationEventArgs(Exception ex)
        {
            Exception = ex;
        }

        public Exception Exception { get; set; }
    }
}
