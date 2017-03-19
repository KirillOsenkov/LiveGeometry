using GuiLabs.Canvas.Renderer;

namespace GuiLabs.Canvas.DrawStyle
{
	// [shapeType("Blocks.AbstractBlockStyle")]
	// This Attribute has only to do with Layouts.
	// For each Layout class, that inherits from AbstractLayout,
	// there is exactly one ShapeStyle object.
	public class ShapeStyle : IShapeStyle
	{
		// the 3 objects from the Canvas library
		private IFillStyleInfo mFillStyleInfo;
		public IFillStyleInfo FillStyleInfo
		{
			get
			{
				return mFillStyleInfo;
			}
			set
			{
				mFillStyleInfo = value;
			}
		}

		private ILineStyleInfo mLineStyleInfo;
		public ILineStyleInfo LineStyleInfo
		{
			get
			{
				return mLineStyleInfo;
			}
			set
			{
				mLineStyleInfo = value;
			}
		}

		private IFontStyleInfo mFontStyleInfo;
		public IFontStyleInfo FontStyleInfo
		{
			get
			{
				return mFontStyleInfo;
			}
			set
			{
				mFontStyleInfo = value;
			}
		}

		/// <summary>
		/// ShapeStyle Constructor initializes the three objects with their default values.
		/// </summary>
		public ShapeStyle()
		{
			FillStyleInfo = RendererSingleton.Instance.DrawOperations.Factory.ProduceNewFillStyleInfo(System.Drawing.Color.White);
			LineStyleInfo = RendererSingleton.Instance.DrawOperations.Factory.ProduceNewLineStyleInfo(System.Drawing.Color.Black, 1);
			FontStyleInfo = RendererSingleton.Instance.DrawOperations.Factory.ProduceNewFontStyleInfo("Courier New", 10, System.Drawing.FontStyle.Regular);
		}

		/// <summary>
		/// Name is the identifier of this concrete style object.
		/// Name binds the object to all objects of its layout class.
		/// </summary>
		private string mName;
		public string Name
		{
			get
			{
				return mName;
			}
			set
			{
				mName = value;
			}
		}

		// The following properties are just shortcuts to the
		// corresponding properties of the 3 basic Canvas objects.

		#region Shortcuts to properties of Canvas objects

		/// <summary>
		/// Shortcut to LineStyle.ForeColor
		/// </summary>
		public System.Drawing.Color LineColor
		{
			get
			{
				return this.LineStyleInfo.ForeColor;
			}
			set
			{
				this.LineStyleInfo.ForeColor = value;
			}
		}

		/// <summary>
		/// Shortcut to FillStyle.FillColor
		/// </summary>
		public System.Drawing.Color FillColor
		{
			get
			{
				return this.FillStyleInfo.FillColor;
			}
			set
			{
				this.FillStyleInfo.FillColor = value;
			}
		}

		/// <summary>
		/// Shortcut to LineStyle.Width
		/// </summary>
		public int LineWidth
		{
			get
			{
				return this.LineStyleInfo.Width;
			}
			set
			{
				this.LineStyleInfo.Width = value;
			}
		}

		/// <summary>
		/// Shortcut to MyFontStyle.Font.Name
		/// </summary>
		public string FontName
		{
			get
			{
				return this.FontStyleInfo.Font.Name;
			}
		}

		/// <summary>
		/// Shortcut to MyFontStyle.Font.Size
		/// </summary>
		public int FontSize
		{
			get
			{
				return (int)this.FontStyleInfo.Font.Size;
			}
		}

		#endregion
	}
}
