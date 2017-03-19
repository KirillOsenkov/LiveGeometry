using System;

namespace GuiLabs.Canvas.Utils
{
	public static class Common
	{
		public static void Swap<T>(ref T Value1, ref T Value2)
		{
			T Temp = Value1;
			Value1 = Value2;
			Value2 = Temp;
		}

		public static T Max<T>(ref T Value1, ref T Value2) 
			where T : System.IComparable
		{
			if (Value1.CompareTo(Value2) < 0)
			{
				return Value2;
			}
			else
			{
				return Value1;
			}
		}

		public static void EnsureGreater<T>(ref T Value, T ComparedWith) 
			where T : System.IComparable
		{
			if (Value.CompareTo(ComparedWith) < 0)
			{
				Value = ComparedWith;
			}
		}

		public static void SwapIfGreater<T>(ref T LValue, ref T RValue)
			where T : IComparable
		{
			if (LValue.CompareTo(RValue) > 0)
			{
				T temp = LValue;
				LValue = RValue;
				RValue = temp;
			}
		}
	}
}
