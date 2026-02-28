/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using System;
namespace NPOI.HSLF.Model.TextProperties
{

    /** 
     * DefInition for the common character text property bitset, which
     *  handles bold/italic/underline etc.
     */
    public class CharFlagsTextProp : BitMaskTextProp
    {
        public static int BOLD_IDX = 0;
        public static int ITALIC_IDX = 1;
        public static int UNDERLINE_IDX = 2;
        public static int SHADOW_IDX = 4;
        public static int STRIKETHROUGH_IDX = 8;
        public static int RELIEF_IDX = 9;
        public static int RESET_NUMBERING_IDX = 10;
        public static int ENABLE_NUMBERING_1_IDX = 11;
        public static int ENABLE_NUMBERING_2_IDX = 12;

        public static String NAME = "char_flags";
        public CharFlagsTextProp():base(2, 0xffff, NAME, new String[] {
				"bold",                 // 0x0001  A bit that specifies whether the characters are bold.
				"italic",               // 0x0002  A bit that specifies whether the characters are italicized.
				"underline",            // 0x0004  A bit that specifies whether the characters are underlined.
				"char_unknown_1",       // 0x0008  Undefined and MUST be ignored.
				"shadow",               // 0x0010  A bit that specifies whether the characters have a shadow effect.
				"fehint",               // 0x0020  A bit that specifies whether characters originated from double-byte input.
				"char_unknown_2",       // 0x0040  Undefined and MUST be ignored.
				"kumi",                 // 0x0080  A bit that specifies whether Kumimoji are used for vertical text.
				"strikethrough",        // 0x0100  Undefined and MUST be ignored.
				"emboss",               // 0x0200  A bit that specifies whether the characters are embossed.
                "char_unknown_3",       // 0x0400  Undefined and MUST be ignored.
                "char_unknown_4",       // 0x0800  Undefined and MUST be ignored.
                "char_unknown_5",       // 0x1000  Undefined and MUST be ignored.
			})
        {
            
        }
    }

}