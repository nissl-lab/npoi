using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Common.UserModel.Fonts
{
    public class FontGroup
    {
        /** type for latin charset (default) - also used for unicode fonts like MS Arial Unicode */
        public static readonly FontGroup LATIN = new FontGroup(1);
        /** type for east asian charsets - usually set as fallback for the latin font, e.g. something like MS Gothic or MS Mincho */
        public static readonly FontGroup EAST_ASIAN = new FontGroup(2);
        /** type for symbol fonts */
        public static readonly FontGroup SYMBOL = new FontGroup(3);
        /** type for complex scripts - see https://msdn.microsoft.com/en-us/library/windows/desktop/dd317698 */
        public static readonly FontGroup COMPLEX_SCRIPT = new FontGroup(4);

        private static Dictionary<int,Range> UCS_RANGES=null;

        static FontGroup()
        {
            UCS_RANGES = new Dictionary<int, Range>();
            UCS_RANGES.Add(0x0000, new Range(0x007F, LATIN));
            UCS_RANGES.Add(0x0080, new Range(0x00A6, LATIN));
            UCS_RANGES.Add(0x00A9, new Range(0x00AF, LATIN));
            UCS_RANGES.Add(0x00B2, new Range(0x00B3, LATIN));
            UCS_RANGES.Add(0x00B5, new Range(0x00D6, LATIN));
            UCS_RANGES.Add(0x00D8, new Range(0x00F6, LATIN));
            UCS_RANGES.Add(0x00F8, new Range(0x058F, LATIN));
            UCS_RANGES.Add(0x0590, new Range(0x074F, COMPLEX_SCRIPT));
            UCS_RANGES.Add(0x0780, new Range(0x07BF, COMPLEX_SCRIPT));
            UCS_RANGES.Add(0x0900, new Range(0x109F, COMPLEX_SCRIPT));
            UCS_RANGES.Add(0x10A0, new Range(0x10FF, LATIN));
            UCS_RANGES.Add(0x1200, new Range(0x137F, LATIN));
            UCS_RANGES.Add(0x13A0, new Range(0x177F, LATIN));
            UCS_RANGES.Add(0x1D00, new Range(0x1D7F, LATIN));
            UCS_RANGES.Add(0x1E00, new Range(0x1FFF, LATIN));
            UCS_RANGES.Add(0x1780, new Range(0x18AF, COMPLEX_SCRIPT));
            UCS_RANGES.Add(0x2000, new Range(0x200B, LATIN));
            UCS_RANGES.Add(0x200C, new Range(0x200F, COMPLEX_SCRIPT));
            // For the quote characters in the range U+2018 - U+201E, use the East Asian font
            // if the text has one of the following language identifiers:
            // ii-CN, ja-JP, ko-KR, zh-CN,zh-HK, zh-MO, zh-SG, zh-TW
            UCS_RANGES.Add(0x2010, new Range(0x2029, LATIN));
            UCS_RANGES.Add(0x202A, new Range(0x202F, COMPLEX_SCRIPT));
            UCS_RANGES.Add(0x2030, new Range(0x2046, LATIN));
            UCS_RANGES.Add(0x204A, new Range(0x245F, LATIN));
            UCS_RANGES.Add(0x2670, new Range(0x2671, COMPLEX_SCRIPT));
            UCS_RANGES.Add(0x27C0, new Range(0x2BFF, LATIN));
            UCS_RANGES.Add(0x3099, new Range(0x309A, EAST_ASIAN));
            UCS_RANGES.Add(0xD835, new Range(0xD835, LATIN));
            UCS_RANGES.Add(0xF000, new Range(0xF0FF, SYMBOL));
            UCS_RANGES.Add(0xFB00, new Range(0xFB17, LATIN));
            UCS_RANGES.Add(0xFB1D, new Range(0xFB4F, COMPLEX_SCRIPT));
            UCS_RANGES.Add(0xFE50, new Range(0xFE6F, LATIN));
            // All others EAST_ASIAN
        }
        private int value = 0;
        protected FontGroup(int value)
        {
            this.value = value;
        }
        /**
 * Try to guess the font group based on the codepoint
 *
 * @param runText the text which font groups are to be analyzed
 * @return the FontGroup
 */
        public static List<FontGroupRange> GetFontGroupRanges(String runText)
        {
            List<FontGroupRange> ttrList = new List<FontGroupRange>();
            if(string.IsNullOrEmpty(runText))
            {
                return ttrList;
            }
            FontGroupRange ttrLast = null;
            int rlen = runText.Length;
            for(int cp, i = 0, charCount; i < rlen; i += charCount)
            {
                cp = runText[i];
                charCount = 1;    //StringInfo.GetNextTextElementLength(runText); TODO:fix this issue

                // don't switch the font group for a few default characters supposedly available in all fonts
                 FontGroup tt;
                if(ttrLast != null)// && " \n\r".IndexOf() > -1)
                {
                    tt = ttrLast.FontGroup;
                }
                else
                {
                    tt = lookup(cp);
                }

                if(ttrLast == null || ttrLast.FontGroup != tt)
                {
                    ttrLast = new FontGroupRange(tt);
                    ttrList.Add(ttrLast);
                }
                ttrLast.IncreaseLength(charCount);
            }
            return ttrList;
        }
        public static FontGroup GetFontGroupFirst(String runText)
        {
            return (runText == null || string.IsNullOrEmpty(runText)) ? LATIN : lookup(runText[0]);
        }

        private static FontGroup lookup(int codepoint)
        {
            // Do a lookup for a match in UCS_RANGES
            Range range =UCS_RANGES.ContainsKey(codepoint) ? UCS_RANGES[codepoint]:null;
            return (range != null && codepoint <= range.Upper) ? range.FontGroup : EAST_ASIAN;
        }
    }

    public class FontGroupRange
    {
        private FontGroup fontGroup;
        private int len = 0;

        public FontGroupRange(FontGroup fontGroup)
        {
            this.fontGroup = fontGroup;
        }

        public int Length
        {
            get
            {
                return len;
            }
        }

        public FontGroup FontGroup
        {
            get
            {
                return fontGroup;
            }
        }

        public void IncreaseLength(int len)
        {
            this.len += len;
        }
    }

    internal class Range
    {
        private int upper;
        private FontGroup fontGroup;
        public Range(int upper, FontGroup fontGroup)
        {
            this.upper = upper;
            this.fontGroup = fontGroup;
        }

        public int Upper
        {
            get
            { 
                return upper; 
            }
        }
        public FontGroup FontGroup { 
            get { return fontGroup; } 
        }
    }
}
