#region Using directives

using System;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
using GuiLabs.Canvas.Utils;

#endregion

namespace GuiLabs.Canvas.DrawStyle
{
	internal class GDIFont : IFontInfo
	{
		#region Constructors

		public GDIFont(String FamilyName, float size)
			: this(FamilyName, size, System.Drawing.FontStyle.Regular)
		{
		}

		public GDIFont(GDIFont ExistingFont)
		{
			this.mName = ExistingFont.Name;
			this.mSize = ExistingFont.Size;
			this.mBold = ExistingFont.Bold;
			this.mItalic = ExistingFont.Italic;
			this.mUnderline = ExistingFont.Underline;
			this.hFont = ExistingFont.hFont;
		}

		public GDIFont(String FamilyName, float size, System.Drawing.FontStyle style)
		{
			Init();

			this.mName = FamilyName;
			this.mSize = (int)size;
			this.mBold = (style & System.Drawing.FontStyle.Bold) != 0;
			this.mItalic = (style & System.Drawing.FontStyle.Italic) != 0;
			this.mUnderline = (style & System.Drawing.FontStyle.Underline) != 0;
			
			AssignHandle();
		}

		#endregion

		#region Init, Exit

		private static bool WasInit = false;

		private void Init()
		{
			if (!WasInit)
			{
				System.Windows.Forms.Application.ApplicationExit += new EventHandler(OnApplicationExit);
				WasInit = true;
			}
		}

		private void OnApplicationExit(object sender, System.EventArgs e)
		{
			if (WasInit)
			{
				System.Windows.Forms.Application.ApplicationExit -= new EventHandler(OnApplicationExit);
				foreach (IntPtr h in FontHandles.Values)
				{
					API.DeleteObject(h);
				}
				WasInit = false;
			}
		}

		#endregion

		private void AssignHandle()
		{
			IntPtr handle = FindHandle(GetSignature());

			if (handle == IntPtr.Zero)
			{
				CreateHandle();
				AddHandle(hFont);
			}
			else
			{
				hFont = handle;
			}
		}

		#region CreateHandle

		private void CreateHandle()
		{
			hFont = API.CreateFont(FontSize(this.Size), 0, 0, 0,
				(this.Bold) ? 700 : 400, (this.Italic) ? 1 : 0,
				(this.Underline) ? 1 : 0, 0, 
				// (int)API.CHARSET.DEFAULT_CHARSET
				1
				, 0, 0, 2, 0, this.Name);
		}

		private int FontSize(int Size)
		{
			const int LOGPIXELSY = 90;

			IntPtr hDC = API.GetDC(API.GetDesktopWindow());
			int result = -API.MulDiv(Size, API.GetDeviceCaps(hDC, LOGPIXELSY), 72);
			API.ReleaseDC(API.GetDesktopWindow(), hDC);

			return result;
		}

		#endregion

		private static Hashtable FontHandles = new Hashtable();

		#region AddHandle, FindHandle

		private void AddHandle(IntPtr h)
		{
			FontHandles[GetSignature()] = h;
		}

		private IntPtr FindHandle(string Signature)
		{
			object o = FontHandles[Signature];

			if (o == null)
			{
				return IntPtr.Zero;
			}

			return (IntPtr)o;
		}

		private string GetSignature()
		{
			StringBuilder s = new StringBuilder();
			s.Append(this.Name);
			s.Append(" ");
			s.Append(this.Size.ToString());
			s.Append(", ");
			s.Append(this.GetFontStyle().ToString());
			return s.ToString();
		}

		#endregion

		#region Properties

		private IntPtr mhFont;
		public IntPtr hFont
		{
			get
			{
				return mhFont;
			}
			set
			{
				mhFont = value;
			}
		}

		private string mName;
		public string Name
		{
			get
			{
				return mName;
			}
		}

		private int mSize;
		public int Size
		{
			get
			{
				return mSize;
			}
		}

		#region Style

		private System.Drawing.FontStyle GetFontStyle()
		{
			System.Drawing.FontStyle result = System.Drawing.FontStyle.Regular;

			if (this.Bold)
			{
				result |= System.Drawing.FontStyle.Bold;
			}

			if (this.Italic)
			{
				result |= System.Drawing.FontStyle.Italic;
			}

			if (this.Underline)
			{
				result |= System.Drawing.FontStyle.Underline;
			}

			return result;
		}

		private bool mBold;
		public bool Bold
		{
			get
			{
				return mBold;
			}
		}

		private bool mItalic;
		public bool Italic
		{
			get
			{
				return mItalic;
			}
		}

		private bool mUnderline;
		public bool Underline
		{
			get
			{
				return mUnderline;
			}
		}

		private Point mSpaceCharSize = new Point();
		public Point SpaceCharSize
		{
			get
			{
				if (mSpaceCharSize.X == 0 || mSpaceCharSize.Y == 0)
				{
					mSpaceCharSize = Renderer.RendererSingleton.DrawOperations.StringSize(" ", this);
				}
				return mSpaceCharSize;
			}
		}

		#endregion

		#endregion
	}
}