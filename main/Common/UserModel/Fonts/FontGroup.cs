using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.Common.UserModel.Fonts
{
	public class FontGroup
	{
		public class FontGroupRange
		{
			public FontGroupEnum FontGroup {get; private set;}
			protected int len = 0;

			public FontGroupRange(FontGroupEnum fontGroup)
			{
				this.FontGroup = fontGroup;
			}

			public int getLength()
			{
				return this.len;
			}

			public void increaseLength(int len)
			{
				this.len += len;
			}
		}

		public class Range
		{
			public int Upper { get; private set; }
			public FontGroupEnum FontGroup { get; private set; }

			public Range(int upper, FontGroupEnum fontGroup)
			{
				this.Upper = upper;
				this.FontGroup = fontGroup;
			}

		}

		private static readonly SortedDictionary<int, Range> UCS_RANGES = new SortedDictionary<int, Range>()
		{
			{ 0x0000,  new Range(0x007F, FontGroupEnum.LATIN) },
			{ 0x0080,  new Range(0x00A6, FontGroupEnum.LATIN) },
			{ 0x00A9,  new Range(0x00AF, FontGroupEnum.LATIN) },
			{ 0x00B2,  new Range(0x00B3, FontGroupEnum.LATIN) },
			{ 0x00B5,  new Range(0x00D6, FontGroupEnum.LATIN) },
			{ 0x00D8,  new Range(0x00F6, FontGroupEnum.LATIN) },
			{ 0x00F8,  new Range(0x058F, FontGroupEnum.LATIN) },
			{ 0x0590,  new Range(0x074F, FontGroupEnum.COMPLEX_SCRIPT) },
			{ 0x0780,  new Range(0x07BF, FontGroupEnum.COMPLEX_SCRIPT) },
			{ 0x0900,  new Range(0x109F, FontGroupEnum.COMPLEX_SCRIPT) },
			{ 0x10A0,  new Range(0x10FF, FontGroupEnum.LATIN) },
			{ 0x1200,  new Range(0x137F, FontGroupEnum.LATIN) },
			{ 0x13A0,  new Range(0x177F, FontGroupEnum.LATIN) },
			{ 0x1D00,  new Range(0x1D7F, FontGroupEnum.LATIN) },
			{ 0x1E00,  new Range(0x1FFF, FontGroupEnum.LATIN) },
			{ 0x1780,  new Range(0x18AF, FontGroupEnum.COMPLEX_SCRIPT) },
			{ 0x2000,  new Range(0x200B, FontGroupEnum.LATIN) },
			{ 0x200C,  new Range(0x200F, FontGroupEnum.COMPLEX_SCRIPT) },
			// For the quote characters in the range U+2018 - U+201E, use the East Asian font
			// if the text has one of the following language identifiers:
			// ii-CN, ja-JP, ko-KR, zh-CN,zh-HK, zh-MO, zh-SG, zh-TW
			{ 0x2010,  new Range(0x2029, FontGroupEnum.LATIN) },
			{ 0x202A,  new Range(0x202F, FontGroupEnum.COMPLEX_SCRIPT) },
			{ 0x2030,  new Range(0x2046, FontGroupEnum.LATIN) },
			{ 0x204A,  new Range(0x245F, FontGroupEnum.LATIN) },
			{ 0x2670,  new Range(0x2671, FontGroupEnum.COMPLEX_SCRIPT) },
			{ 0x27C0,  new Range(0x2BFF, FontGroupEnum.LATIN) },
			{ 0x3099,  new Range(0x309A, FontGroupEnum.EAST_ASIAN) },
			{ 0xD835,  new Range(0xD835, FontGroupEnum.LATIN) },
			{ 0xF000,  new Range(0xF0FF, FontGroupEnum.SYMBOL) },
			{ 0xFB00,  new Range(0xFB17, FontGroupEnum.LATIN) },
			{ 0xFB1D,  new Range(0xFB4F, FontGroupEnum.COMPLEX_SCRIPT) },
			{ 0xFE50,  new Range(0xFE6F, FontGroupEnum.LATIN) }
			// All others EAST_ASIAN
		};


		/**
		* Try to guess the font group based on the codepoint
		 *
		 * @param runText the text which font groups are to be analyzed
		 * @return the FontGroup
		 */
		public static List<FontGroupRange> GetFontGroupRanges(string runText)
		{
			List<FontGroupRange> ttrList = new List<FontGroupRange>();
			if (runText == null || string.IsNullOrEmpty(runText))
			{
				return ttrList;
			}
			FontGroupRange ttrLast = null;
			int rlen = runText.Length;
			for (int cp, i = 0, charCount; i < rlen; i += charCount)
			{
				cp = (int)runText[i];
				charCount = Character.CharCount(cp);

				// don't switch the font group for a few default characters supposedly available in all fonts
				FontGroupEnum tt;
				if (ttrLast != null && " \n\r".IndexOf(runText[i]) > -1)
				{
					tt = ttrLast.FontGroup;
				}
				else
				{
					tt = lookup(cp);
				}

				if (ttrLast == null || ttrLast.FontGroup != tt)
				{
					ttrLast = new FontGroupRange(tt);
					ttrList.Add(ttrLast);
				}
				ttrLast.increaseLength(charCount);
			}
			return ttrList;
		}

		public static FontGroupEnum getFontGroupFirst(String runText)
		{
			return (runText == null || string.IsNullOrEmpty(runText)) ? FontGroupEnum.LATIN : lookup((int)runText[0]);
		}

		private static FontGroupEnum lookup(int codepoint)
		{
			// Do a lookup for a match in UCS_RANGES
			//var floorEntry = UCS_RANGES.Keys.ToList().Where(k => k <= codepoint).Max();
			KeyValuePair<int, Range> entry = UCS_RANGES.LastOrDefault(k => k.Key <= codepoint);
			Range range = (entry.Value != null) ? entry.Value : null;
			return (range != null && codepoint <= range.Upper) ? range.FontGroup : FontGroupEnum.EAST_ASIAN;
		}
	}
}
