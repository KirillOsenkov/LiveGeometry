using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DynamicGeometry
{
    public partial class DrawingControl
    {
        public Command CommandUndo { get; set; }
        public Command CommandRedo { get; set; }
    }
}
