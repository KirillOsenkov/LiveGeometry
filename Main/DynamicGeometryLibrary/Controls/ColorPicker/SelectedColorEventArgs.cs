using System;
using System.Windows.Media;

namespace SilverlightContrib.Controls
{
    /// <summary>
    /// Delegate for the SelectedColorChanged event.
    /// </summary>
    /// <param name="sender">The object instance that fired the event.</param>
    /// <param name="e">The selected color event arguments for the event.</param>
    public delegate void SelectedColorChangedHandler(object sender, SelectedColorEventArgs e);

    /// <summary>
    /// Delegate for the SelectedColorChanging event.
    /// </summary>
    /// <param name="sender">The object instance that fired the event.</param>
    /// <param name="e">The selected color event arguments for the event.</param>
    public delegate void SelectedColorChangingHandler(object sender, SelectedColorEventArgs e);


    /// <summary>
    /// Event data for the SelectedColorChanged event.
    /// </summary>
    public class SelectedColorEventArgs : EventArgs
    {
        /// <summary>
        /// The currently selected color.
        /// </summary>
        public readonly Color SelectedColor;

        /// <summary>
        /// Create a new instance of the SelectedColorEventArgs class.
        /// </summary>
        /// <param name="color">The currently selected color.</param>
        public SelectedColorEventArgs(Color color)
        {
            this.SelectedColor = color;
        }


    }
}
