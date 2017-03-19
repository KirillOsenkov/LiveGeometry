namespace GuiLabs.Canvas.Utils
{
	public static class Colors
	{
		public static System.Drawing.Color ScaleColor(System.Drawing.Color SourceColor, float ScaleFactor)
		{
			int Win32Color = System.Drawing.ColorTranslator.ToWin32(SourceColor);
			float r = R(Win32Color) * ScaleFactor;
			float g = G(Win32Color) * ScaleFactor;
			float b = B(Win32Color) * ScaleFactor;

			if (r < 0) r = 0;
			if (g < 0) g = 0;
			if (b < 0) b = 0;

			if (r > 255) r = 255;
			if (g > 255) g = 255;
			if (b > 255) b = 255;

			return System.Drawing.Color.FromArgb((int)r, (int)g, (int)b);
		}

		public static int ScaleColor(int SourceColor, float ScaleFactor)
		{
			float r = R(SourceColor) * ScaleFactor;
			float g = G(SourceColor) * ScaleFactor;
			float b = B(SourceColor) * ScaleFactor;

			if (r < 0) r = 0;
			if (g < 0) g = 0;
			if (b < 0) b = 0;

			if (r > 255) r = 255;
			if (g > 255) g = 255;
			if (b > 255) b = 255;

			return RGB((int)r, (int)g, (int)b);
		}

		public static int R(int SourceColor)
		{
			return SourceColor & 0xFF;
		}

		public static int G(int SourceColor)
		{
			return (SourceColor & 0xFF00) >> 8;
		}

		public static int B(int SourceColor)
		{
			return (SourceColor & 0xFF0000) >> 16;
		}

		public static int RGB(int R, int G, int B)
		{
			return (B << 16) | (G << 8) | R;
		}
	}
}