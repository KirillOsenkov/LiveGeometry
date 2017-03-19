namespace DynamicGeometry
{
    public class FreePointCreator : Behavior
    {
        protected override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.RightButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                foreach (var item in Drawing.Figures)
                {
                    item.Recalculate();
                }
            }

            CreatePointAtCurrentPosition(e);
        }

        public override string Icon
        {
            get
            {
                return "resources/bitmaps/geometry%20toolbar/dgwpoint.bmp";
            }
        }
    }
}