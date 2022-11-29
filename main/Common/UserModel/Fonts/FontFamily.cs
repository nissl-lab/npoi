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
namespace NPOI.Common.UserModel.Fonts
{
	using System;
	using System.Collections.Generic;
	using System.Text;

	/**
 * A property of a font that describes its general appearance.
 * 
 * @since POI 3.17-beta2
 */
	public class FontFamily
	{
		private int nativeId;
		private FontFamilyEnum native;
		private FontFamily(int nativeId)
		{
			this.nativeId = nativeId;
			this.native = (FontFamilyEnum)nativeId;
		}

		public int getFlag()
		{
			return nativeId;
		}

		public static FontFamily ValueOf(int nativeId)
		{
			foreach (var item in Enum.GetValues(typeof(FontFamilyEnum)))
			{
				if (nativeId == (int)item)
				{
					return new FontFamily(nativeId);
				}
			}
			return null;
		}

		/**
	     * Get FontFamily from combined native id
	     *
	     * @param pitchAndFamily The PitchFamily to decode.
	     *
	     * @return The resulting FontFamily
	     */
		public static FontFamily ValueOfPitchFamily(byte pitchAndFamily)
		{
			return ValueOf(pitchAndFamily >> 4);
		}
	}
}