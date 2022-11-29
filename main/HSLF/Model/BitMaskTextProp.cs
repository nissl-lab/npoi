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
using System.Text;
using System;
using System.Collections.Generic;
using NPOI.Util;
using NPOI.Common;

namespace NPOI.HSLF.Model
{
	/**
 * Definition of a special kind of property of some text, or its
 *  paragraph. For these properties, a flag in the "contains" header
 *  field tells you the data property family will exist. The value
 *  of the property is itself a mask, encoding several different
 *  (but related) properties
 */
	public abstract class BitMaskTextProp: TextProp, IDuplicatable<BitMaskTextProp>
	{
		private string[] subPropNames;
		private int[] subPropMasks;
		private bool[] subPropMatches;

		/** Fetch the list of the names of the sub properties */
		public string[] GetSubPropNames() { return subPropNames; }
		/** Fetch the list of if the sub properties match or not */
		public bool[] GetSubPropMatches() { return subPropMatches; }


		public BitMaskTextProp(BitMaskTextProp other)
			:base(other)
		{
			subPropNames = (string[])((other.subPropNames == null) ? null : other.subPropNames.Clone());
			subPropMasks = (int[])((other.subPropMasks == null) ? null : other.subPropMasks.Clone());

			// The old clone implementation didn't carry over matches, but keep everything else as it was
			// this is failing unit tests
			// subPropMatches = (other.subPropMatches == null) ? null : new bool[other.subPropMatches.length];
			subPropMatches = (bool[])((other.subPropMatches == null) ? null : other.subPropMatches.Clone());
		}


		protected BitMaskTextProp(int sizeOfDataBlock, int maskInHeader, string overallName, params string[] subPropNames)
			:base(sizeOfDataBlock, maskInHeader, overallName)
		{
			this.subPropNames = subPropNames;
			subPropMasks = new int[subPropNames.Length];
			subPropMatches = new bool[subPropNames.Length];

			int LSB = Integer.LowestOneBit(maskInHeader);

			// Initialise the masks list
			for (int i = 0; i < subPropMasks.Length; i++)
			{
				subPropMasks[i] = (LSB << i);
			}
		}

		/**
		 * Calculate mask from the subPropMatches.
		 */
		//@Override
		public new int GetWriteMask()
		{
			/*
			 * The dataValue can't be taken as a mask, as sometimes certain properties
			 * are explicitly set to false, i.e. the mask says the property is defined
			 * but in the actually nibble the property is set to false
			 */
			int mask = 0, i = 0;
			foreach (int subMask in subPropMasks)
			{
				if (subPropMatches[i++]) mask |= subMask;
			}
			return mask;
		}

		/**
		 * Sets the write mask, i.e. which defines the text properties to be considered
		 *
		 * @param writeMask the mask, bit values outside the property mask range will be ignored
		 */
		public void SetWriteMask(int writeMask)
		{
			int i = 0;
			foreach (int subMask in subPropMasks)
			{
				subPropMatches[i++] = ((writeMask & subMask) != 0);
			}
		}

		/**
		 * Return the text property value.
		 * Clears all bits of the value, which are marked as unset.
		 *
		 * @return the text property value.
		 */
		//@Override
		public new int GetValue()
		{
			return maskValue(base.GetValue());
		}

		private int maskValue(int pVal)
		{
			int val = pVal, i = 0;
			foreach (int mask in subPropMasks)
			{
				if (!subPropMatches[i++])
				{
					val &= ~mask;
				}
			}
			return val;
		}

		/**
		 * Set the value of the text property, and recompute the sub
		 * properties based on it, i.e. all unset subvalues will be cleared.
		 * Use {@link #setSubValue(bool, int)} to explicitly set subvalues to {@code false}.
		 */
		//@Override
		public new void SetValue(int val)
		{
			base.SetValue(val);

			// Figure out the values of the sub properties
			int i = 0;
			foreach (int mask in subPropMasks)
			{
				subPropMatches[i++] = ((val & mask) != 0);
			}
		}

		/**
		 * Convenience method to set a value with mask, without splitting it into the subvalues
		 */
		public void SetValueWithMask(int val, int writeMask)
		{
			SetWriteMask(writeMask);
			base.SetValue(maskValue(val));
			if (val != base.GetValue())
			{
				//LOG.atWarn().log("Style properties of '{}' don't match mask - output will be sanitized", getName());
				//LOG.atDebug().log(()-> {
				//	StringBuilder sb = new StringBuilder("The following style attributes of the '")
				//			.append(getName()).append("' property will be ignored:\n");
				//	int i = 0;
				//	for (int mask : subPropMasks)
				//	{
				//		if (!subPropMatches[i] && (val & mask) != 0)
				//		{
				//			sb.append(subPropNames[i]).append(",");
				//		}
				//		i++;
				//	}
				//	return new SimpleMessage(sb);
				//});
			}
		}

		/**
		 * Fetch the true/false status of the subproperty with the given index
		 */
		public bool GetSubValue(int idx)
		{
			return subPropMatches[idx] && ((base.GetValue() & subPropMasks[idx]) != 0);
		}

		/**
		 * Set the true/false status of the subproperty with the given index
		 */
		public void SetSubValue(bool value, int idx)
		{
			subPropMatches[idx] = true;
			int newVal = base.GetValue();
			if (value)
			{
				newVal |= subPropMasks[idx];
			}
			else
			{
				newVal &= ~subPropMasks[idx];
			}
			base.SetValue(newVal);
		}

		/**
		 * @return an identical copy of this, i.e. also the subPropMatches are copied
		 */
		public BitMaskTextProp CloneAll()
		{
			BitMaskTextProp bmtp = Copy();
			if (subPropMatches != null)
			{
				Array.Copy(subPropMatches, 0, bmtp.subPropMatches, 0, subPropMatches.Length);
			}
			return bmtp;
		}

		//@Override
		public new IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
				"base", ()=> base.GetGenericProperties<T>(),
				"flags", () => GenericRecordUtil.GetBitsAsString(GetValue, subPropMasks, subPropNames)
			);
		}

		//@Override
		public new abstract BitMaskTextProp Copy();
	}
}