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
using NPOI.HSLF.Record;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.Model
{
	/**
 * Definition of a property of some text, or its paragraph. Defines
 * how to find out if it's present (via the mask on the paragraph or
 * character "contains" header field), how long the value of it is,
 * and how to get and set the value.
 *
 * As the exact form of these (such as mask value, size of data
 *  block etc) is different for StyleTextProps and
 *  TxMasterTextProps, the definitions of the standard
 *  TextProps is stored in the different record classes
 */
	public class TextProp : GenericRecord, IDuplicatable<TextProp>
	{
		private int sizeOfDataBlock; // Number of bytes the data part uses
		private string propName;
		private int dataValue;
		private int maskInHeader;

		/**
		 * Generate the definition of a given type of text property.
		 */
		public TextProp(int sizeOfDataBlock, int maskInHeader, string propName)
		{
			this.sizeOfDataBlock = sizeOfDataBlock;
			this.maskInHeader = maskInHeader;
			this.propName = propName;
			this.dataValue = 0;
		}

		/**
		 * Clones the property
		 */
		public TextProp(TextProp other)
		{
			this.sizeOfDataBlock = other.sizeOfDataBlock;
			this.maskInHeader = other.maskInHeader;
			this.propName = other.propName;
			this.dataValue = other.dataValue;
		}

		/**
		 * Name of the text property
		 */
		public string GetName() { return propName; }

		/**
		 * Size of the data section of the text property (2 or 4 bytes)
		 */
		public int GetSize() { return sizeOfDataBlock; }

		/**
		 * Mask in the paragraph or character "contains" header field
		 *  that indicates that this text property is present.
		 */
		public int GetMask() { return maskInHeader; }
		/**
		 * Get the mask that's used at write time. Only differs from
		 *  the result of getMask() for the mask based properties
		 */
		public int GetWriteMask() { return GetMask(); }

		/**
		 * Fetch the value of the text property (meaning is specific to
		 *  each different kind of text property)
		 */
		public int GetValue() { return dataValue; }

		/**
		 * Set the value of the text property.
		 */
		public void SetValue(int val) { dataValue = val; }

		/**
		 * Clone, eg when you want to actually make use of one of these.
		 */
		//@Override
		public TextProp Copy()
		{
			return new TextProp(this);
		}

		//@Override
		public override int GetHashCode()
		{
			//Objects.hash(dataValue, maskInHeader, propName, sizeOfDataBlock);
			return Tuple.Create(dataValue, maskInHeader, propName, sizeOfDataBlock).GetHashCode();
		}

		//@Override
	public override bool Equals(object obj)
	{
		if (this == obj)
		{
			return true;
		}
		if (obj == null)
		{
			return false;
		}
		if (this.GetType() != obj.GetType())
		{
			return false;
		}
		TextProp other = (TextProp)obj;
		if (dataValue != other.dataValue)
		{
			return false;
		}
		if (maskInHeader != other.maskInHeader)
		{
			return false;
		}
		if (propName == null)
		{
			if (other.propName != null)
			{
				return false;
			}
		}
		else if (!propName.Equals(other.propName))
		{
			return false;
		}
		if (sizeOfDataBlock != other.sizeOfDataBlock)
		{
			return false;
		}
		return true;
	}

	//@Override
	public override string ToString()
	{
		int len;
		switch (GetSize())
		{
			case 1: len = 4; break;
			case 2: len = 6; break;
			default: len = 10; break;
		}
			//String.format(Locale.ROOT, "%s = %d (%0#" + len + "X mask / %d bytes)", getName(), getValue(), getMask(), getSize());
			return String.Format(" {0} = {1:d}({2}#" + len + "X mask / {3:d} bytes)", GetName(), GetValue(), Convert.ToString(GetMask(),8), GetSize());
	}

	//@Override
	public IDictionary<string, Func<T>> GetGenericProperties<T>()
	{
		return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
			"sizeOfDataBlock", () => GetSize(),
			"propName", () => GetName(),
			"dataValue", () => GetValue(),
			"maskInHeader", ()=> GetMask()
		);
	}

		public RecordTypes GetGenericRecordType()
		{
			throw new NotImplementedException();
		}

		public IList<GenericRecord> GetGenericChildren()
		{
			throw new NotImplementedException();
		}
	}
}