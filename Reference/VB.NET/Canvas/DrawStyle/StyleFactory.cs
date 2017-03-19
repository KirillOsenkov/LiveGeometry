using System.Collections.Generic;

namespace GuiLabs.Canvas.DrawStyle
{
	using StyleDic = Dictionary<string, IShapeStyle>;

	/// <summary>
	/// Globally visible singleton collection of all Styles for all Layouts.
	/// </summary>
	public class StyleFactory
	{
		protected StyleFactory()
		{

		}

		#region Singleton Instances

		private static StyleFactory mStylesInstance;
		public static StyleFactory Styles
		{
			get 
			{
				if (mStylesInstance == null)
				{
					mStylesInstance = new StyleFactory();
				}
				return mStylesInstance; 
			}
			set 
			{ 
				mStylesInstance = value; 
			}
		}

		private static StyleFactory mSelectedStylesInstance;
		public static StyleFactory SelectedStyles
		{
			get 
			{
				if (mSelectedStylesInstance == null)
				{
					mSelectedStylesInstance = new StyleFactory();
				}
				return mSelectedStylesInstance;
			}
			set 
			{
				mSelectedStylesInstance = value; 
			}
		}

		#endregion

		/// <summary>
		/// (String -> ShapeStyle) hashtable. 
		/// Not visible from outside.
		/// </summary>
		protected StyleDic mStyleList = new StyleDic();
		protected StyleDic StyleList
		{
			get
			{
				return mStyleList;
			}
		}

		/// <param name="shapeType">Typically Layout.StyleName</param>
		/// <returns>ShapeStyle with the specified name, if it exists, null otherwise.
		/// </returns>
		public IShapeStyle GetStyle(string shapeType)
		{
			// It'child OK when there are no styles, GetStyle will just return null.
			/*if (!IsInit)
			{
				Init();
			}
			*/
			IShapeStyle result = null;
			StyleList.TryGetValue(shapeType, out result);
			return result;
		}

		public void Add(IShapeStyle NewStyle)
		{
			if (NewStyle == null)
				return;
			StyleList.Add(NewStyle.Name, NewStyle);
		}
	}
}
