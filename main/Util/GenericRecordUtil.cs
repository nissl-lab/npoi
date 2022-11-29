using NPOI.Common.UserModel.Fonts;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Util
{
	public class GenericRecordUtil
	{
		public static Dictionary<string, Func<T>> GetGenericProperties<T>(string v, Func<T> value)
		{
			return new Dictionary<string, Func<T>>() { { v, value } };
		}

		public static IDictionary<string, Func<object>> GetGenericProperties(string v1, Func<object> sup1, string v2, Func<object> sup2)
		{
			return GetGenericProperties(v1, sup1, v2, sup2, null, null, null, null, null, null, null, null);
		}

		public static IDictionary<string, Func<object>> GetGenericProperties(string v1, Func<object> sup1, string v2, Func<object> sup2, string v3, Func<object> sup3, string v4, Func<object> sup4)
		{ 
			return GetGenericProperties(v1, sup1, v2, sup2,v3, sup3, v4, sup4, null, null, null, null);
		}

		public static IDictionary<string, Func<object>> GetGenericProperties(string v1, Func<object> sup1, string v2, Func<object> sup2, string v3, Func<object> sup3, string v4, Func<object> sup4, string v5, Func<object> sup5, string v6, Func<object> sup6)
		{
			string[] vals = new string[] { v1,v2,v3,v4,v5,v6};
			Func<object>[] sups = new Func<object>[] { sup1, sup2, sup3, sup4, sup5, sup6 };
			IDictionary<string, Func<object>> result = new Dictionary<string, Func<object>>();
			for (int i=0; i< vals.Length && vals[i] != null; i++)
			{
				if ("base".Equals(vals[i]))
				{
					object baseMap = sups[i]();
					if (baseMap is IDictionary<string, Func<T>>)
					{
						result.Concat((IDictionary<string, Func<object>>)baseMap);
					}
				}
				else
				{
					result.Add(vals[i], sups[i]);
				}
			}
			// TODO
			return result;
		}


		public static Func<AnnotatedFlag> GetBitsAsString(Func<int> flags, BitField[] masks, string[] names)
		{
			int[] iMasks = masks.Select(m => m.GetMask()).ToArray();
			return ()=> new AnnotatedFlag(flags, iMasks, names, false);
		}

		public static Func<AnnotatedFlag> GetBitsAsString(Func<int> flags, int[] masks, string[] names)
		{
			return ()=> new AnnotatedFlag(flags, masks, names, false);
		}

		

		public class AnnotatedFlag
		{
			private Func<int> value;
			private Dictionary<int, String> masks = new Dictionary<int, string>();
			private bool exactMatch;

			public AnnotatedFlag(Func<int> value, int[] masks, String[] names, bool exactMatch)
			{
				//assert(masks.length == names.length);
				if (masks.Length == names.Length)
				{
					this.value = value;
					this.exactMatch = exactMatch;
					for (int i = 0; i < masks.Length; i++)
					{
						this.masks.Add(masks[i], names[i]);
					}
				}
			}

			public Func<int> GetValue()
			{
				return value;
			}

			public String GetDescription()
			{
				int val = (int)value();
				return String.Join(" | ", masks[val].ToArray());
			}

			private bool Match(int val, int mask)
			{
				return exactMatch ? (val == mask) : ((val & mask) == mask);
			}
		}
	}


}
