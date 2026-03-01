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


using NPOI.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters;
using System.Text;

namespace NPOI.Common.UserModel.Fonts
{

    /// <summary>
    /// Text runs can contain characters which will be handled (if configured) by a different font,
    /// because the default font (latin) doesn't contain corresponding glyphs.
    /// </summary>
    /// 
    /// @see <a href="https://blogs.msdn.microsoft.com/officeinteroperability/2013/04/22/office-open-xml-themes-schemes-and-fonts/">Office Open XML Themes, Schemes, and Fonts</a>
    /// <remarks>
    /// @since POI 3.17-beta2
    /// </remarks>
    public enum FontGroup
    {
        /** type for latin charset (default) - also used for unicode fonts like MS Arial Unicode */
        LATIN,
        /** type for east asian charsets - usually Set as fallback for the latin font, e.g. something like MS Gothic or MS Mincho */
        EAST_ASIAN,
        /** type for symbol fonts */
        SYMBOL,
        /** type for complex scripts - see https://msdn.microsoft.com/en-us/library/windows/desktop/dd317698 */
        COMPLEX_SCRIPT
    }
    internal class Range
    {
        internal int upper;
        internal FontGroup fontGroup;
        public Range(int upper, FontGroup fontGroup)
        {
            this.upper = upper;
            this.fontGroup = fontGroup;
        }
    }
    public class FontGroupRange
    {
        internal int len;
        internal FontGroup fontGroup;
        public int GetLength()
        {
            return len;
        }
        public FontGroup GetFontGroup()
        {
            return fontGroup;
        }
    }
    public class FontGroupExtension
    {
        private static Dictionary<int,Range> UCS_RANGES;
        private static SortedSet<int> sortedKeys = new();
        static FontGroupExtension()
        {
            UCS_RANGES = new Dictionary<int, Range> {
                { 0x0000, new Range(0x007F, FontGroup.LATIN) },
                { 0x0080, new Range(0x00A6, FontGroup.LATIN) },
                { 0x00A9, new Range(0x00AF, FontGroup.LATIN) },
                { 0x00B2, new Range(0x00B3, FontGroup.LATIN) },
                { 0x00B5, new Range(0x00D6, FontGroup.LATIN) },
                { 0x00D8, new Range(0x00F6, FontGroup.LATIN) },
                { 0x00F8, new Range(0x058F, FontGroup.LATIN) },
                { 0x0590, new Range(0x074F, FontGroup.COMPLEX_SCRIPT) },
                { 0x0780, new Range(0x07BF, FontGroup.COMPLEX_SCRIPT) },
                { 0x0900, new Range(0x109F, FontGroup.COMPLEX_SCRIPT) },
                { 0x10A0, new Range(0x10FF, FontGroup.LATIN) },
                { 0x1200, new Range(0x137F, FontGroup.LATIN) },
                { 0x13A0, new Range(0x177F, FontGroup.LATIN) },
                { 0x1D00, new Range(0x1D7F, FontGroup.LATIN) },
                { 0x1E00, new Range(0x1FFF, FontGroup.LATIN) },
                { 0x1780, new Range(0x18AF, FontGroup.COMPLEX_SCRIPT) },
                { 0x2000, new Range(0x200B, FontGroup.LATIN) },
                { 0x200C, new Range(0x200F, FontGroup.COMPLEX_SCRIPT) },
                // For the quote characters in the range U+2018 - U+201E, use the East Asian font
                // if the text has one of the following language identifiers:
                // ii-CN, ja-JP, ko-KR, zh-CN,zh-HK, zh-MO, zh-SG, zh-TW
                { 0x2010, new Range(0x2029, FontGroup.LATIN) },
                { 0x202A, new Range(0x202F, FontGroup.COMPLEX_SCRIPT) },
                { 0x2030, new Range(0x2046, FontGroup.LATIN) },
                { 0x204A, new Range(0x245F, FontGroup.LATIN) },
                { 0x2670, new Range(0x2671, FontGroup.COMPLEX_SCRIPT) },
                { 0x27C0, new Range(0x2BFF, FontGroup.LATIN) },
                { 0x3099, new Range(0x309A, FontGroup.EAST_ASIAN) },
                { 0xD835, new Range(0xD835, FontGroup.LATIN) },
                { 0xF000, new Range(0xF0FF, FontGroup.SYMBOL) },
                { 0xFB00, new Range(0xFB17, FontGroup.LATIN) },
                { 0xFB1D, new Range(0xFB4F, FontGroup.COMPLEX_SCRIPT) },
                { 0xFE50, new Range(0xFE6F, FontGroup.LATIN) }
            };
            // All others EAST_ASIAN

            foreach(var kv in UCS_RANGES)
                sortedKeys.Add(kv.Key);
        }


        /**
         * Try to guess the font group based on the codepoint
         *
         * @param runText the text which font groups are to be analyzed
         * @return the FontGroup
         */
        public static List<FontGroupRange> GetFontGroupRanges(string runText)
        {
            List<FontGroupRange> ttrList = new ();
            if(string.IsNullOrEmpty(runText))
            {
                return ttrList;
            }
            FontGroupRange ttrLast = null;
            int rlen = runText.Length;
            for(int cp, i = 0, charCount; i < rlen; i += charCount)
            {
                cp = runText.CodePointAt(i);
                charCount = StringUtil.CharCount(cp);

                // don't switch the font group for a few default characters supposedly available in all fonts
                FontGroup tt;
                if(ttrLast != null && " \n\r".IndexOf((char)cp) > -1)
                {
                    tt = ttrLast.fontGroup;
                }
                else
                {
                    tt = Lookup(cp);
                }

                if(ttrLast == null || ttrLast.fontGroup != tt)
                {
                    ttrLast = new FontGroupRange();
                    ttrLast.fontGroup = tt;
                    ttrList.Add(ttrLast);
                }
                ttrLast.len += charCount;
            }
            return ttrList;
        }

        public static FontGroup GetFontGroupFirst(string runText)
        {
            return string.IsNullOrEmpty(runText) ? FontGroup.LATIN : Lookup(runText.CodePointAt(0));
        }

        private static FontGroup Lookup(int codepoint)
        {
            // Do a lookup for a match in UCS_RANGES
            int key = sortedKeys.GetViewBetween(0, codepoint).Max;
            //KeyValuePair<int,Range> entry = UCS_RANGES.FloorEntry(codepoint);
            //Range range = entry.Key > 0 ? entry.Value : null;
            UCS_RANGES.TryGetValue(key, out var range);
            return (range != null && codepoint <= range.upper) ? range.fontGroup : FontGroup.EAST_ASIAN;
        }
    }
}


