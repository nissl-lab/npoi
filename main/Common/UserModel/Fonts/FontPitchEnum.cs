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
	 * A property of a font that describes the pitch, of the characters.
	 * 
	 * @since POI 3.17-beta2
	 */
	public enum FontPitchEnum
	{
	    /**
	     * The default pitch, which is implementation-dependent.
	     */
	    DEFAULT = 0x00,
	    /**
	     * A fixed pitch, which means that all the characters in the font occupy the same
	     * width when output in a string.
	     */
	    FIXED = 0x01,
	    /**
	     * A variable pitch, which means that the characters in the font occupy widths
	     * that are proportional to the actual widths of the glyphs when output in a string. For example,
	     * the "i" and space characters usually have much smaller widths than a "W" or "O" character.
	     */
	    VARIABLE = 0x02
	}
}
