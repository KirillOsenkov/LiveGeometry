using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GuiLabs.Canvas.Utils
{
	internal class Log
	{
		protected Log()
		{
		}

		private static Log mInstance = null;
		public static Log Instance
		{
			get
			{
				if (mInstance == null)
				{
					mInstance = new Log();
				}
				return mInstance;
			}
		}

		private string mLogFileName = "GuiLabsCanvasLog.txt";
		public string LogFileName
		{
			get
			{
				return mLogFileName;
			}
			set
			{
				mLogFileName = value;
			}
		}

		public void WriteWarning(string Text)
		{
			using (StreamWriter writer = new StreamWriter(LogFileName, true))
			{
				writer.WriteLine("Warning: " + Text);
			}
		}
	}
}
