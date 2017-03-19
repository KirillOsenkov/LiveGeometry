using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Collections.Generic;

namespace DynamicGeometry
{
    public partial class DrawingControl
    {
        public bool ConstructionInProgress { get; set; }

        public void Undo()
        {
            if (ConstructionInProgress)
            {
                Drawing.Behavior.Restart();
                return;
            }
            try
            {
                Drawing.ActionManager.Undo();
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
            UpdateUndoRedo();
            Drawing.RaiseDisplayProperties(null);
        }

        public void Redo()
        {
            try
            {
                Drawing.ActionManager.Redo();
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
            UpdateUndoRedo();
            Drawing.RaiseDisplayProperties(null);
        }

        private void ActionManager_CollectionChanged(object sender, EventArgs e)
        {
            UpdateUndoRedo();
        }

        private void UpdateUndoRedo()
        {
            try
            {
                CommandUndo.Enabled = Drawing.ActionManager.CanUndo || ConstructionInProgress;
                CommandRedo.Enabled = Drawing.ActionManager.CanRedo;
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        private void mCurrentDrawing_ConstructionStepStarted(object sender, Drawing.ConstructionStepStartedEventArgs e)
        {
            ConstructionInProgress = true;
            UpdateUndoRedo();
        }

        private void mCurrentDrawing_ConstructionStepComplete(object sender, Drawing.ConstructionStepCompleteEventArgs args)
        {
            if (args.ConstructionComplete)
            {
                ConstructionInProgress = false;
                UpdateUndoRedo();
                Drawing.ClearStatus();
            }
            else
            {
                Drawing.RaiseDisplayProperties(Drawing.Behavior.PropertyBag);
                CommandRedo.Enabled = false;
                Drawing.RaiseStatusNotification(Drawing.Behavior.ConstructionHintText(args));
            }
        }

        protected virtual void mCurrentDrawing_FiguresBeingAdded(object sender, Drawing.UIAFEventArgs args)
        {
            // Do nothing.  I override this in Tabula.  - D.H.
        }

    }
}
