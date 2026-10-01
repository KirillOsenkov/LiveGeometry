using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Media;

namespace DynamicGeometry
{
    public partial class DrawingControl : Canvas
    {
        public event EventHandler ReadyForInteraction;
        public event Action<Drawing> DrawingDetach = delegate { };
        public event Action<Drawing> DrawingAttach = delegate { };

        private Drawing mCurrentDrawing;
        public Drawing Drawing
        {
            get
            {
                return mCurrentDrawing;
            }
            set
            {
                if (mCurrentDrawing != null)
                {
                    DrawingDetach(mCurrentDrawing);
                    mCurrentDrawing.ActionManager.CollectionChanged -= ActionManager_CollectionChanged;
                    mCurrentDrawing.ConstructionStepStarted -= mCurrentDrawing_ConstructionStepStarted;
                    mCurrentDrawing.ConstructionStepComplete -= mCurrentDrawing_ConstructionStepComplete;
                    mCurrentDrawing.DocumentOpenRequested -= mCurrentDrawing_DocumentOpenRequested;
                    mCurrentDrawing.UserIsAddingFigures -= mCurrentDrawing_FiguresBeingAdded;
                    mCurrentDrawing.Canvas = null;

                    // A construction left half way goes with its drawing: taking the canvas
                    // away stops the tool, but nobody hears it say so any more. Still "in
                    // progress" for the next drawing, Undo only restarted the tool and Redo
                    // did nothing, until some tool finished a figure.
                    ConstructionInProgress = false;
                }
                mCurrentDrawing = value;
                if (mCurrentDrawing != null)
                {
                    mCurrentDrawing.ActionManager.CollectionChanged += ActionManager_CollectionChanged;
                    mCurrentDrawing.ConstructionStepStarted += mCurrentDrawing_ConstructionStepStarted;
                    mCurrentDrawing.ConstructionStepComplete += mCurrentDrawing_ConstructionStepComplete;
                    mCurrentDrawing.DocumentOpenRequested += mCurrentDrawing_DocumentOpenRequested;
                    mCurrentDrawing.UserIsAddingFigures += mCurrentDrawing_FiguresBeingAdded;
                    DrawingAttach(mCurrentDrawing);
                    mCurrentDrawing.SetDefaultBehavior();
                }
                UpdateUndoRedo();
            }
        }

        public DrawingControl()
        {
            // the paper, until a drawing paints its own over it: bound below the value a
            // drawing sets, or a switch of theme would paint the theme's paper over that again
            this.Bind(BackgroundProperty, this.GetResourceObservable(nameof(AppTheme.Paper)), BindingPriority.Style);
            // while on screen; hidden behind the gallery page, the drawing catches up when the
            // editor shows again (whoever shows it calls Drawing.RefreshThemeIfStale: Avalonia
            // tells nobody when IsEffectivelyVisible changes)
            AppTheme.CurrentChanged += () => RefreshTheme(colorsChanged: false);
            AppTheme.ColorsChanged += () => RefreshTheme(colorsChanged: true);
            this.SizeChanged += DrawingControl_SizeChanged;

            CommandUndo = new Command(Undo, null, "Undo", "Drawing");
            CommandRedo = new Command(Redo, null, "Redo", "Drawing");
        }

        void RefreshTheme(bool colorsChanged)
        {
            if (IsEffectivelyVisible)
            {
                Drawing?.RefreshTheme(colorsChanged);
            }
        }

        private void HandleException(Exception ex)
        {
            Drawing.RaiseError(this, ex);
        }

        public virtual void DrawingControl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (this.Drawing == null)
            {
                this.SizeChanged -= DrawingControl_SizeChanged;
                this.Drawing = new Drawing(this);
                if (ReadyForInteraction != null)
                {
                    ReadyForInteraction(this, null);
                }
            }
        }

        public virtual void Clear()
        {
            // the new drawing brings the theme's paper back with it
            Drawing = new Drawing(this);
        }

    }
}
