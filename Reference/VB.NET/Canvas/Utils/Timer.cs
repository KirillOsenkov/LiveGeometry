using System.Runtime.InteropServices;

namespace GuiLabs.Canvas.Utils
{
	public class Timer
	{
		[DllImport("kernel32")]
		private static extern bool QueryPerformanceCounter(out long lpPerformanceCount);
		[DllImport("kernel32")]
		private static extern bool QueryPerformanceFrequency(out long lpFrequency);

		long _Freq, _Start, _Finish;
		double msFreq;

		public Timer()
		{
			QueryPerformanceFrequency(out _Freq);
			msFreq = _Freq / 1000;
		}

		public void Start()
		{
			QueryPerformanceCounter(out _Start);
		}

		public void Stop()
		{
			QueryPerformanceCounter(out _Finish);
		}

		long t;
		public double Milliseconds()
		{
			QueryPerformanceCounter(out t);
			return t / msFreq;
		}

		public double TimeElapsed
		{
			get
			{
				return (double)(_Finish - _Start) / _Freq;
			}
		}
	}
}
