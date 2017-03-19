using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas.DrawStyle
{
	public interface IFontInfo
	{
		string Name
		{
			get;
		}
		int Size
		{
			get;
		}
		bool Bold
		{
			get;
		}
		bool Italic
		{
			get;
		}
		bool Underline
		{
			get;
		}
		Point SpaceCharSize
		{
			get;
		}
	}
}