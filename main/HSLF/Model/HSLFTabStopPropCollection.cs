/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
using NPOI.SL.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.HSLF.Model
{
	/**
 * Container for tabstop lists
 */
	public class HSLFTabStopPropCollection : TextProp
	{
		public static string NAME = "tabStops";

		private List<HSLFTabStop> tabStops = new List<HSLFTabStop>();

		public HSLFTabStopPropCollection()
			:base(0, 0x100000, NAME)
		{
			
		}

		public HSLFTabStopPropCollection(HSLFTabStopPropCollection other)
			:base(other)
		{
			foreach (var item in other.tabStops)
			{
				tabStops.Add(item.Copy());
			}
		}

		/**
		 * Parses the tabstops from TxMasterStyle record
		 *
		 * @param data the data stream
		 * @param offset the offset within the data
		 */
		public void ParseProperty(byte[] data, int offset)
		{
			tabStops.Concat(ReadTabStops(new LittleEndianByteArrayInputStream(data, offset)));
		}

		public static List<HSLFTabStop> ReadTabStops(ILittleEndianInput lei)
		{
			int count = lei.ReadUShort();
			List<HSLFTabStop> tabs = new List<HSLFTabStop>();
			for (int i = 0; i < count; i++)
			{
				int position = lei.ReadShort();
				TabStopType type = TabStopType.fromNativeId(lei.ReadShort());
				tabs.Add(new HSLFTabStop(position, type));
			}
			return tabs;
		}


		public void WriteProperty(OutputStream _out)
		{
			WriteTabStops(new LittleEndianOutputStream(_out), tabStops);
		}

		public static void WriteTabStops(ILittleEndianOutput leo, List<HSLFTabStop> tabStops)
		{
			int count = tabStops.Count;
			leo.WriteShort(count);
			foreach (HSLFTabStop ts in tabStops)
			{
				leo.WriteShort(ts.GetPosition());
				leo.WriteShort(ts.GetType().nativeId);
			}

		}

		//@Override
		public new int GetValue() { return tabStops.Count; }


		//@Override
		public new int GetSize()
		{
			return LittleEndianConsts.SHORT_SIZE + tabStops.Count * LittleEndianConsts.INT_SIZE;
		}

		public List<HSLFTabStop> GetTabStops()
		{
			return tabStops;
		}

		public void ClearTabs()
		{
			tabStops.Clear();
		}

		public void AddTabStop(HSLFTabStop ts)
		{
			tabStops.Add(ts);
		}

		//@Override
		public HSLFTabStopPropCollection Copy()
		{
			return new HSLFTabStopPropCollection(this);
		}

		//@Override
		public override int GetHashCode()
		{
			return Tuple.Create(base.GetHashCode(), tabStops).GetHashCode();
		}

		//@Override
		public override bool Equals(Object obj)
		{
			if (this == obj)
			{
				return true;
			}
			if (!(obj is HSLFTabStopPropCollection)) {
				return false;
			}
			HSLFTabStopPropCollection other = (HSLFTabStopPropCollection)obj;
			if (!base.Equals(other))
			{
				return false;
			}

			return tabStops.Equals(other.tabStops);
		}

		//@Override
		public override string ToString()
		{
			StringBuilder sb = new StringBuilder(base.ToString());
			sb.Append(" [ ");
			bool isFirst = true;
			foreach (HSLFTabStop tabStop in tabStops)
			{
				if (!isFirst)
				{
					sb.Append(", ");
				}
				sb.Append(tabStop.GetType());
				sb.Append(" @ ");
				sb.Append(tabStop.GetPosition());
				isFirst = false;
			}
			sb.Append(" ]");

			return sb.ToString();
		}


		//@Override
		public new IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
				"base", () => (T)base.GetGenericProperties<T>(),
				"tabStops", GetTabStops
			);
		}
	}
}