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
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.Model
{
	/**
 * Definition for the font alignment property.
 */
	public class FontAlignmentProp : TextProp
	{
		public const string NAME = "fontAlign";
        public const int BASELINE = 0;
		public const int TOP = 1;
		public const int CENTER = 2;
		public const int BOTTOM = 3;

		public FontAlignmentProp()
			:base(2, 0x10000, NAME)
		{
		}

		public FontAlignmentProp(FontAlignmentProp other)
			:base(other)
		{
		}

		public FontAlign GetFontAlign()
		{
			switch (GetValue())
			{
				default:
					return FontAlign.AUTO;
				case BASELINE:
					return FontAlign.BASELINE;
				case TOP:
					return FontAlign.TOP;
				case CENTER:
					return FontAlign.CENTER;
				case BOTTOM:
					return FontAlign.BOTTOM;
			}
		}

		//@Override
		public new IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
				"base", ()=> (T)base.GetGenericProperties<T>(),
				"fontAlign", ()=> GetFontAlign()
			);
		}

		//@Override
		public new FontAlignmentProp Copy()
		{
			return new FontAlignmentProp(this);
		}
	}
}