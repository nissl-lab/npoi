using NPOI.SS.Formula.Functions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Text;

namespace NPOI.SL.UserModel
{
    public class AutoNumberingScheme : IEnumerable<AutoNumberingScheme>
    {
        /** Lowercase alphabetic character enclosed in parentheses. Example: (a), (b), (c), ... */
        public static readonly AutoNumberingScheme alphaLcParenBoth = new AutoNumberingScheme(0x0008, 1, "alphaLcParenBoth");
        /** Uppercase alphabetic character enclosed in parentheses. Example: (A), (B), (C), ... */
        public static readonly AutoNumberingScheme alphaUcParenBoth = new AutoNumberingScheme(0x000A, 2, "alphaUcParenBoth");
        /** Lowercase alphabetic character followed by a closing parenthesis. Example: a), b), c), ... */
        public static readonly AutoNumberingScheme alphaLcParenRight = new AutoNumberingScheme(0x0009, 3, "alphaLcParenRight");
        /** Uppercase alphabetic character followed by a closing parenthesis. Example: A), B), C), ... */
        public static readonly AutoNumberingScheme alphaUcParenRight = new AutoNumberingScheme(0x000B, 4, "alphaUcParenRight");
        /** Lowercase Latin character followed by a period. Example: a., b., c., ... */
        public static readonly AutoNumberingScheme alphaLcPeriod = new AutoNumberingScheme(0x0000, 5, "alphaLcPeriod");
        /** Uppercase Latin character followed by a period. Example: A., B., C., ... */
        public static readonly AutoNumberingScheme alphaUcPeriod = new AutoNumberingScheme(0x0001, 6, "alphaUcPeriod");
        /** Arabic numeral enclosed in parentheses. Example: (1), (2), (3), ... */
        public static readonly AutoNumberingScheme arabicParenBoth = new AutoNumberingScheme(0x000C, 7, "arabicParenBoth");
        /** Arabic numeral followed by a closing parenthesis. Example: 1), 2), 3), ... */
        public static readonly AutoNumberingScheme arabicParenRight = new AutoNumberingScheme(0x0002, 8, "arabicParenRight");
        /** Arabic numeral followed by a period. Example: 1., 2., 3., ... */
        public static readonly AutoNumberingScheme arabicPeriod = new AutoNumberingScheme(0x0003, 9, "arabicPeriod");
        /** Arabic numeral. Example: 1, 2, 3, ... */
        public static readonly AutoNumberingScheme arabicPlain = new AutoNumberingScheme(0x000D, 10, "arabicPlain");
        /** Lowercase Roman numeral enclosed in parentheses. Example: (i), (ii), (iii), ... */
        public static readonly AutoNumberingScheme romanLcParenBoth = new AutoNumberingScheme(0x0004, 11, "romanLcParenBoth");
        /** Uppercase Roman numeral enclosed in parentheses. Example: (I), (II), (III), ... */
        public static readonly AutoNumberingScheme romanUcParenBoth = new AutoNumberingScheme(0x000E, 12, "romanUcParenBoth");
        /** Lowercase Roman numeral followed by a closing parenthesis. Example: i), ii), iii), ... */
        public static readonly AutoNumberingScheme romanLcParenRight = new AutoNumberingScheme(0x0005, 13, "romanLcParenRight");
        /** Uppercase Roman numeral followed by a closing parenthesis. Example: I), II), III), .... */
        public static readonly AutoNumberingScheme romanUcParenRight = new AutoNumberingScheme(0x000F, 14, "romanUcParenRight");
        /** Lowercase Roman numeral followed by a period. Example: i., ii., iii., ... */
        public static readonly AutoNumberingScheme romanLcPeriod = new AutoNumberingScheme(0x0006, 15, "romanLcPeriod");
        /** Uppercase Roman numeral followed by a period. Example: I., II., III., ... */
        public static readonly AutoNumberingScheme romanUcPeriod = new AutoNumberingScheme(0x0007, 16, "romanUcPeriod");
        /** Double byte circle numbers. */
        public static readonly AutoNumberingScheme circleNumDbPlain = new AutoNumberingScheme(0x0012, 17, "circleNumDbPlain");
        /** Wingdings black circle numbers. */
        public static readonly AutoNumberingScheme circleNumWdBlackPlain = new AutoNumberingScheme(0x0014, 18, "circleNumWdBlackPlain");
        /** Wingdings white circle numbers. */
        public static readonly AutoNumberingScheme circleNumWdWhitePlain = new AutoNumberingScheme(0x0013, 19, "circleNumWdWhitePlain");
        /** Double-byte Arabic numbers with double-byte period. */
        public static readonly AutoNumberingScheme arabicDbPeriod = new AutoNumberingScheme(0x001D, 20, "arabicDbPeriod");
        /** Double-byte Arabic numbers. */
        public static readonly AutoNumberingScheme arabicDbPlain = new AutoNumberingScheme(0x001C, 21, "arabicDbPlain");
        /** Simplified Chinese with single-byte period. */
        public static readonly AutoNumberingScheme ea1ChsPeriod = new AutoNumberingScheme(0x0011, 22, "ea1ChsPeriod");
        /** Simplified Chinese. */
        public static readonly AutoNumberingScheme ea1ChsPlain = new AutoNumberingScheme(0x0010, 23, "ea1ChsPlain");
        /** Traditional Chinese with single-byte period. */
        public static readonly AutoNumberingScheme ea1ChtPeriod = new AutoNumberingScheme(0x0015, 24, "ea1ChtPeriod");
        /** Traditional Chinese. */
        public static readonly AutoNumberingScheme ea1ChtPlain = new AutoNumberingScheme(0x0014, 25, "ea1ChtPlain");
        /** Japanese with double-byte period. */
        public static readonly AutoNumberingScheme ea1JpnChsDbPeriod = new AutoNumberingScheme(0x0026, 26, "ea1JpnChsDbPeriod");
        /** Japanese/Korean. */
        public static readonly AutoNumberingScheme ea1JpnKorPlain = new AutoNumberingScheme(0x001A, 27, "ea1JpnKorPlain");
        /** Japanese/Korean with single-byte period. */
        public static readonly AutoNumberingScheme ea1JpnKorPeriod = new AutoNumberingScheme(0x001B, 28, "ea1JpnKorPeriod");
        /** Bidi Arabic 1 (AraAlpha) with ANSI minus symbol. */
        public static readonly AutoNumberingScheme arabic1Minus = new AutoNumberingScheme(0x0017, 29, "arabic1Minus");
        /** Bidi Arabic 2 (AraAbjad) with ANSI minus symbol. */
        public static readonly AutoNumberingScheme arabic2Minus = new AutoNumberingScheme(0x0018, 30, "arabic2Minus");
        /** Bidi Hebrew 2 with ANSI minus symbol. */
        public static readonly AutoNumberingScheme hebrew2Minus = new AutoNumberingScheme(0x0019, 31, "hebrew2Minus");
        /** Thai alphabetic character followed by a period. */
        public static readonly AutoNumberingScheme thaiAlphaPeriod = new AutoNumberingScheme(0x001E, 32, "thaiAlphaPeriod");
        /** Thai alphabetic character followed by a closing parenthesis. */
        public static readonly AutoNumberingScheme thaiAlphaParenRight = new AutoNumberingScheme(0x001F, 33, "thaiAlphaParenRight");
        /** Thai alphabetic character enclosed by parentheses. */
        public static readonly AutoNumberingScheme thaiAlphaParenBoth = new AutoNumberingScheme(0x0020, 34, "thaiAlphaParenBoth");
        /** Thai numeral followed by a period. */
        public static readonly AutoNumberingScheme thaiNumPeriod = new AutoNumberingScheme(0x0021, 35, "thaiNumPeriod");
        /** Thai numeral followed by a closing parenthesis. */
        public static readonly AutoNumberingScheme thaiNumParenRight = new AutoNumberingScheme(0x0022, 36, "thaiNumParenRight");
        /** Thai numeral enclosed in parentheses. */
        public static readonly AutoNumberingScheme thaiNumParenBoth = new AutoNumberingScheme(0x0023, 37, "thaiNumParenBoth");
        /** Hindi alphabetic character followed by a period. */
        public static readonly AutoNumberingScheme hindiAlphaPeriod = new AutoNumberingScheme(0x0024, 38, "hindiAlphaPeriod");
        /** Hindi numeric character followed by a period. */
        public static readonly AutoNumberingScheme hindiNumPeriod = new AutoNumberingScheme(0x0025, 39, "hindiNumPeriod");
        /** Hindi numeric character followed by a closing parenthesis. */
        public static readonly AutoNumberingScheme hindiNumParenRight = new AutoNumberingScheme(0x0027, 40, "hindiNumParenRight");
        /** Hindi alphabetic character followed by a period. */
        public static readonly AutoNumberingScheme hindiAlpha1Period = new AutoNumberingScheme(0x0027, 41, "hindiAlpha1Period");


        private static IEnumerable<AutoNumberingScheme> Values
        {
            get {
                /** Lowercase alphabetic character enclosed in parentheses. Example: (a), (b), (c), ... */
                yield return alphaLcParenBoth;
                /** Uppercase alphabetic character enclosed in parentheses. Example: (A), (B), (C), ... */
                yield return alphaUcParenBoth;
                /** Lowercase alphabetic character followed by a closing parenthesis. Example: a), b), c), ... */
                yield return alphaLcParenRight;
                /** Uppercase alphabetic character followed by a closing parenthesis. Example: A), B), C), ... */
                yield return alphaUcParenRight;
                /** Lowercase Latin character followed by a period. Example: a., b., c., ... */
                yield return alphaLcPeriod;
                /** Uppercase Latin character followed by a period. Example: A., B., C., ... */
                yield return alphaUcPeriod;
                /** Arabic numeral enclosed in parentheses. Example: (1), (2), (3), ... */
                yield return arabicParenBoth;
                /** Arabic numeral followed by a closing parenthesis. Example: 1), 2), 3), ... */
                yield return arabicParenRight;
                /** Arabic numeral followed by a period. Example: 1., 2., 3., ... */
                yield return arabicPeriod;
                /** Arabic numeral. Example: 1, 2, 3, ... */
                yield return arabicPlain;
                /** Lowercase Roman numeral enclosed in parentheses. Example: (i), (ii), (iii), ... */
                yield return romanLcParenBoth;
                /** Uppercase Roman numeral enclosed in parentheses. Example: (I), (II), (III), ... */
                yield return romanUcParenBoth;
                /** Lowercase Roman numeral followed by a closing parenthesis. Example: i), ii), iii), ... */
                yield return romanLcParenRight;
                /** Uppercase Roman numeral followed by a closing parenthesis. Example: I), II), III), .... */
                yield return romanUcParenRight;
                /** Lowercase Roman numeral followed by a period. Example: i., ii., iii., ... */
                yield return romanLcPeriod;
                /** Uppercase Roman numeral followed by a period. Example: I., II., III., ... */
                yield return romanUcPeriod;
                /** Double byte circle numbers. */
                yield return circleNumDbPlain;
                /** Wingdings black circle numbers. */
                yield return circleNumWdBlackPlain;
                /** Wingdings white circle numbers. */
                yield return circleNumWdWhitePlain;
                /** Double-byte Arabic numbers with double-byte period. */
                yield return arabicDbPeriod;
                /** Double-byte Arabic numbers. */
                yield return arabicDbPlain;
                /** Simplified Chinese with single-byte period. */
                yield return ea1ChsPeriod;
                /** Simplified Chinese. */
                yield return ea1ChsPlain;
                /** Traditional Chinese with single-byte period. */
                yield return ea1ChtPeriod;
                /** Traditional Chinese. */
                yield return ea1ChtPlain;
                /** Japanese with double-byte period. */
                yield return ea1JpnChsDbPeriod;
                /** Japanese/Korean. */
                yield return ea1JpnKorPlain;
                /** Japanese/Korean with single-byte period. */
                yield return ea1JpnKorPeriod;
                /** Bidi Arabic 1 (AraAlpha) with ANSI minus symbol. */
                yield return arabic1Minus;
                /** Bidi Arabic 2 (AraAbjad) with ANSI minus symbol. */
                yield return arabic2Minus;
                /** Bidi Hebrew 2 with ANSI minus symbol. */
                yield return hebrew2Minus;
                /** Thai alphabetic character followed by a period. */
                yield return thaiAlphaPeriod;
                /** Thai alphabetic character followed by a closing parenthesis. */
                yield return thaiAlphaParenRight;
                /** Thai alphabetic character enclosed by parentheses. */
                yield return thaiAlphaParenBoth;
                /** Thai numeral followed by a period. */
                yield return thaiNumPeriod;
                /** Thai numeral followed by a closing parenthesis. */
                yield return thaiNumParenRight;
                /** Thai numeral enclosed in parentheses. */
                yield return thaiNumParenBoth;
                /** Hindi alphabetic character followed by a period. */
                yield return hindiAlphaPeriod;
                /** Hindi numeric character followed by a period. */
                yield return hindiNumPeriod;
                /** Hindi numeric character followed by a closing parenthesis. */
                yield return hindiNumParenRight;
                /** Hindi alphabetic character followed by a period. */
                yield return hindiAlpha1Period;
            }
        }

        public IEnumerator<AutoNumberingScheme> GetEnumerator()
        {
            return (IEnumerator<AutoNumberingScheme>)Values;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }


        public int nativeId, ooxmlId;
        public string name = string.Empty;

        AutoNumberingScheme(int nativeId, int ooxmlId, string name)
        {
            this.nativeId = nativeId;
            this.ooxmlId = ooxmlId;
            this.name = name;
        }

        public static AutoNumberingScheme ForNativeID(int nativeId)
        {
            foreach (AutoNumberingScheme ans in Values)
            {
                if (ans.nativeId == nativeId) return ans;
            }
            return null;
        }

        public static AutoNumberingScheme forOoxmlID(int ooxmlId)
        {
            foreach (AutoNumberingScheme ans in Values)
            {
                if (ans.ooxmlId == ooxmlId) return ans;
            }
            return null;
        }

        public String GetDescription()
        {
            switch (this.name)
            {
				case "alphaLcPeriod": return "Lowercase Latin character followed by a period. Example: a., b., c., ...";
				case "alphaUcPeriod": return "Uppercase Latin character followed by a period. Example: A., B., C., ...";
				case "arabicParenRight": return "Arabic numeral followed by a closing parenthesis. Example: 1), 2), 3), ...";
				case "arabicPeriod": return "Arabic numeral followed by a period. Example: 1., 2., 3., ...";
				case "romanLcParenBoth": return "Lowercase Roman numeral enclosed in parentheses. Example: (i), (ii), (iii), ...";
				case "romanLcParenRight": return "Lowercase Roman numeral followed by a closing parenthesis. Example: i), ii), iii), ...";
				case "romanLcPeriod": return "Lowercase Roman numeral followed by a period. Example: i., ii., iii., ...";
				case "romanUcPeriod": return "Uppercase Roman numeral followed by a period. Example: I., II., III., ...";
				case "alphaLcParenBoth": return "Lowercase alphabetic character enclosed in parentheses. Example: (a), (b), (c), ...";
				case "alphaLcParenRight": return "Lowercase alphabetic character followed by a closing parenthesis. Example: a), b), c), ...";
				case "alphaUcParenBoth": return "Uppercase alphabetic character enclosed in parentheses. Example: (A), (B), (C), ...";
				case "alphaUcParenRight": return "Uppercase alphabetic character followed by a closing parenthesis. Example: A), B), C), ...";
				case "arabicParenBoth": return "Arabic numeral enclosed in parentheses. Example: (1), (2), (3), ...";
				case "arabicPlain": return "Arabic numeral. Example: 1, 2, 3, ...";
				case "romanUcParenBoth": return "Uppercase Roman numeral enclosed in parentheses. Example: (I), (II), (III), ...";
				case "romanUcParenRight": return "Uppercase Roman numeral followed by a closing parenthesis. Example: I), II), III), ...";
				case "ea1ChsPlain": return "Simplified Chinese.";
				case "ea1ChsPeriod": return "Simplified Chinese with single-byte period.";
				case "circleNumDbPlain": return "Double byte circle numbers.";
				case "circleNumWdWhitePlain": return "Wingdings white circle numbers.";
				case "circleNumWdBlackPlain": return "Wingdings black circle numbers.";
				case "ea1ChtPlain": return "Traditional Chinese.";
				case "ea1ChtPeriod": return "Traditional Chinese with single-byte period.";
				case "arabic1Minus": return "Bidi Arabic 1 (AraAlpha) with ANSI minus symbol.";
				case "arabic2Minus": return "Bidi Arabic 2 (AraAbjad) with ANSI minus symbol.";
				case "hebrew2Minus": return "Bidi Hebrew 2 with ANSI minus symbol.";
				case "ea1JpnKorPlain": return "Japanese/Korean.";
				case "ea1JpnKorPeriod": return "Japanese/Korean with single-byte period.";
				case "arabicDbPlain": return "Double-byte Arabic numbers.";
				case "arabicDbPeriod": return "Double-byte Arabic numbers with double-byte period.";
				case "thaiAlphaPeriod": return "Thai alphabetic character followed by a period.";
				case "thaiAlphaParenRight": return "Thai alphabetic character followed by a closing parenthesis.";
				case "thaiAlphaParenBoth": return "Thai alphabetic character enclosed by parentheses.";
				case "thaiNumPeriod": return "Thai numeral followed by a period.";
				case "thaiNumParenRight": return "Thai numeral followed by a closing parenthesis.";
				case "thaiNumParenBoth": return "Thai numeral enclosed in parentheses.";
				case "hindiAlphaPeriod": return "Hindi alphabetic character followed by a period.";
				case "hindiNumPeriod": return "Hindi numeric character followed by a period.";
				case "ea1JpnChsDbPeriod": return "Japanese with double-byte period.";
				case "hindiNumParenRight": return "Hindi numeric character followed by a closing parenthesis.";
				case "hindiAlpha1Period": return "Hindi alphabetic character followed by a period.";
				default: return "Unknown Numbered Scheme";
			}
        }

        public String Format(int value)
        {
            String index = FormatIndex(value);
            String cased = FormatCase(index);
            return FormatSeperator(cased);
        }

        private String FormatSeperator(String cased)
        {
            String name = this.name.ToLower(CultureInfo.InvariantCulture);
            if (name.Contains("plain")) return cased;
            if (name.Contains("parenright")) return cased + ")";
            if (name.Contains("parenboth")) return "(" + cased + ")";
            if (name.Contains("period")) return cased + ".";
            if (name.Contains("minus")) return cased + "-"; // ???
            return cased;
        }

        private String FormatCase(String index)
        {
            String name = this.name.ToLower(CultureInfo.InvariantCulture);
			if (name.Contains("lc")) return index.ToLower(CultureInfo.InvariantCulture);
            if (name.Contains("uc")) return index.ToUpper(CultureInfo.InvariantCulture);
			return index;
        }

        private static String ARABIC_LIST = "0123456789";
        private static String ALPHA_LIST = "abcdefghijklmnopqrstuvwxyz";
        private static String WINGDINGS_WHITE_LIST =
                "\u0080\u0081\u0082\u0083\u0084\u0085\u0086\u0087\u0088\u0089";
        private static String WINGDINGS_BLACK_LIST =
                "\u008B\u008C\u008D\u008E\u008F\u0090\u0091\u0092\u0093\u0094";
        private static String CIRCLE_DB_LIST =
                "\u2776\u2777\u2778\u2779\u277A\u277B\u277C\u277D\u277E";

        private String FormatIndex(int value)
        {
            String name = this.name.ToLower(CultureInfo.InvariantCulture);
			if (name.StartsWith("roman"))
            {
                return FormatRomanIndex(value);
            }
            else if (name.StartsWith("arabic") && !name.Contains("db"))
            {
                return GetIndexedList(value, ARABIC_LIST, false);
            }
            else if (name.StartsWith("alpha"))
            {
                return GetIndexedList(value, ALPHA_LIST, true);
            }
            else if (name.Contains("wdwhite"))
            {
                return (value == 10) ? "\u008A"
                    : GetIndexedList(value, WINGDINGS_WHITE_LIST, false);
            }
            else if (name.Contains("wdblack"))
            {
                return (value == 10) ? "\u0095"
                    : GetIndexedList(value, WINGDINGS_BLACK_LIST, false);
            }
            else if (name.Contains("numdb"))
            {
                return (value == 10) ? "\u277F"
                    : GetIndexedList(value, CIRCLE_DB_LIST, true);
            }
            else
            {
                return "?";
            }
        }

        private static String GetIndexedList(int val, String list, bool oneBased)
        {
            StringBuilder sb = new StringBuilder();
            AddIndexedChar(val, list, oneBased, sb);
            return sb.ToString();
        }

        private static void AddIndexedChar(int val, String list, bool oneBased, StringBuilder sb)
        {
            if (oneBased) val -= 1;
            int len = list.Length;
            if (val >= len)
            {
                AddIndexedChar(val / len, list, oneBased, sb);
            }
            sb.Append(list[(val % len)]);
        }


        private String FormatRomanIndex(int value)
        {
            //M (1000), CM (900), D (500), CD (400), C (100), XC (90), L (50), XL (40), X (10), IX (9), V (5), IV (4) and I (1).
            int[] VALUES = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] ROMAN = new string[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            List<(string a, string b)> conciseList = new List<(string a, string b)>()
            {
                ("XLV", "VL"), //45
                ("XCV", "VC"), //95
                ("CDL", "LD"), //450
                ("CML", "LM"), //950
                ("CMVC", "LMVL"), //995
                ("CDXC", "LDXL"), //490
                ("CDVC", "LDVL"), //495
                ("CMXC", "LMXL"), //990
                ("XCIX", "VCIV"), //99
                ("XLIX", "VLIV"), //49
                ("XLIX", "IL"), //49
                ("XCIX", "IC"), //99
                ("CDXC", "XD"), //490
                ("CDVC", "XDV"), //495
                ("CDIC", "XDIX"), //499
                ("LMVL", "XMV"), //995
                ("CMIC", "XMIX"), //999
                ("CMXC", "XM"), // 990
                ("XDV", "VD"),  //495
                ("XDIX", "VDIV"), //499
                ("XMV", "VM"), // 995
                ("XMIX", "VMIV"), //999
                ("VDIV", "ID"), //499
                ("VMIV", "IM") //999

            };

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < 13; i++)
            {
                while (value >= VALUES[i])
                {
                    value -= VALUES[i];
                    sb.Append(ROMAN[i]);
                }
            }
            string result = sb.ToString();
            foreach (var cc in conciseList)
            {
                result = result.Replace(cc.a, cc.b);
            }
            return result;
        }
    }
}