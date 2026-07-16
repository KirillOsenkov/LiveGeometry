using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;
using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry
{
    public partial class DrawingControl
    {
        public Command CommandUndo { get; set; }
        public Command CommandRedo { get; set; }
    }
}
