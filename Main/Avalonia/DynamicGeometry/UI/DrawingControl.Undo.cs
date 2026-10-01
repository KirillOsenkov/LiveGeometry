using System;

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

            // a drag under way: its moves are not in the history yet, and undoing what is
            // would leave them on top of a state they were not made in
            if (Drawing.IsRecordingTransaction)
            {
                return;
            }

            try
            {
                Drawing.ActionManager.Undo();
            }
            catch (Exception ex)
            {
                ReleaseStuckAction();
                HandleException(ex);
            }
            UpdateUndoRedo();
            CommandToolButton.UpdateToggles();
            Drawing.RaiseDisplayProperties(null);
        }

        public void Redo()
        {
            // not in the middle of a construction or a drag: what is redone would land among
            // the figures being made (its points took the names of the ones under way)
            if (ConstructionInProgress || Drawing.IsRecordingTransaction)
            {
                return;
            }

            try
            {
                Drawing.ActionManager.Redo();
            }
            catch (Exception ex)
            {
                ReleaseStuckAction();
                HandleException(ex);
            }
            UpdateUndoRedo();
            CommandToolButton.UpdateToggles();
            Drawing.RaiseDisplayProperties(null);
        }

        /// <summary>
        /// The undo library marks the action it is undoing or redoing and takes the mark off
        /// when it is done - but not when the action threw, and with the mark left on it
        /// refuses every action from then on ("currently running or undoing an action"): one
        /// undo that failed, and nothing more could be done to the drawing. The mark has no
        /// public setter.
        /// </summary>
        void ReleaseStuckAction()
        {
            var manager = Drawing.ActionManager;
            if (manager.ActionIsExecuting)
            {
                typeof(GuiLabs.Undo.ActionManager)
                    .GetProperty(nameof(GuiLabs.Undo.ActionManager.CurrentAction))
                    .SetValue(manager, null);
            }
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
                // a tool whose panel belongs to a step of the construction has none now
                Drawing.RaiseDisplayProperties(Drawing.Behavior?.PropertyBag);
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
