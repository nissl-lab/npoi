using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.Util
{
	public class Integer
	{
		public static int LowestOneBit(int number)
		{
			if (number == 0) return 0;
			return (int)Math.Pow(2, Convert.ToString(number, 2).Reverse().ToList().IndexOf('1'));
		}

		public static int HighestOneBit(int number)
		{
			if (number == 0) return 0;
			var _bin = Convert.ToString(number, 2).ToList();
			return (int)Math.Pow(2, _bin.Count - 1 - _bin.IndexOf('1'));
		}
	}
}
