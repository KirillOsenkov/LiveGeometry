namespace DynamicGeometry
{
    public class Dragger : Behavior
    {
        public override string Icon
        {
            get
            {
                return "resources/bitmaps/geometry%20toolbar/dgwpointer.bmp";
            }
        }

        private FreePoint dragging = null;

        protected override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var found = Drawing.Figures.HitTest(e.GetPosition(Drawing.Parent));
            dragging = found as FreePoint;

            //var result = VisualTreeHelper.HitTest(
            //    Parent, 
            //    e.GetPosition(Parent));

            //FrameworkElement ell = result.VisualHit as FrameworkElement;
            //if (ell != null)
            //{
            //    dragging = ell;
            //    return;
            //}
        }

        protected override void MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (dragging != null)
            {
                dragging.MoveTo(e.GetPosition(Parent));
                Drawing.Figures.Recalculate();
            }
        }

        protected override void MouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            dragging = null;
        }
    }
}