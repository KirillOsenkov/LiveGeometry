using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

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
            this.Background = new SolidColorBrush(Colors.White);
            this.SizeChanged += DrawingControl_SizeChanged;

            CommandUndo = new Command(Undo, null, "Undo", "Drawing");
            CommandRedo = new Command(Redo, null, "Redo", "Drawing");
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
            Drawing = new Drawing(this);
        }

    }
}
