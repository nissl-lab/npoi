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
using NPOI.Common;
using NPOI.Common.UserModel;
using NPOI.SL.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.Model
{
	public class HSLFTabStop: TabStop, IDuplicatable<HSLFTabStop>, GenericRecord
	{
		/**
     * A signed integer that specifies an offset, in master units, of the tab stop.
     *
     * If the TextPFException record that contains this TabStop structure also contains a
     * leftMargin, then the value of position is relative to the left margin of the paragraph;
     * otherwise, the value is relative to the left side of the paragraph.
     *
     * If a TextRuler record contains this TabStop structure, the value is relative to the
     * left side of the text ruler.
     */
		private int position;

		/**
		 * A enumeration that specifies how text aligns at the tab stop.
		 */
		private TabStopType type;

		public HSLFTabStop(int position, TabStopType type)
		{
			this.position = position;
			this.type = type;
		}

		public HSLFTabStop(HSLFTabStop other)
		{
			position = other.position;
			type = other.type;
		}

		public int GetPosition()
		{
			return position;
		}

		public void SetPosition(int position)
		{
			this.position = position;
		}

		//@Override
		public double GetPositionInPoints()
		{
			return Units.MasterToPoints(GetPosition());
		}

		//@Override
		public void SetPositionInPoints(double points)
		{
			SetPosition(Units.PointsToMaster(points));
		}

		//@Override
		public TabStopType GetType()
		{
			return type;
		}

		//@Override
		public void SetType(TabStopType type)
		{
			this.type = type;
		}

		//@Override
		public HSLFTabStop Copy()
		{
			return new HSLFTabStop(this);
		}

		//@Override
		public override int GetHashCode()
		{
			return Tuple.Create(position, type).GetHashCode();
		}
		
		//@Override
		public override bool Equals(Object obj)
		{
			if (this == obj)
			{
				return true;
			}
			if (!(obj is HSLFTabStop)) {
				return false;
			}
			HSLFTabStop other = (HSLFTabStop)obj;
			if (position != other.position)
			{
				return false;
			}
			if (type != other.type)
			{
				return false;
			}
			return true;
		}

		//@Override
		public override string ToString()
		{
			return type + " @ " + position;
		}

		//@Override
		public IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
				"type", GetType,
				"position", ()=> GetPosition()
			);
		}
	}
}