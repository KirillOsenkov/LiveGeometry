namespace GuiLabs.Canvas.DrawStyle
{
    public interface IShapeStyle
    {
		ILineStyleInfo LineStyleInfo 
		{
			get;
			set;
		}

		IFillStyleInfo FillStyleInfo
		{
			get;
			set;
		}

		IFontStyleInfo FontStyleInfo
		{
			get;
			set;
		}

		string Name
		{
			get;
			set;
		}

		// TODO: TO THINK: Do we need these properties?
		System.Drawing.Color LineColor
		{
			get;
			set;
		}
		System.Drawing.Color FillColor
		{
			get;
			set;
		}
		int LineWidth
		{
			get;
			set;
		}
		string FontName
		{
			get;
		}
		int FontSize
		{
			get;
		}
	}
}
