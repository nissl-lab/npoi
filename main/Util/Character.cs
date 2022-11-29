using System.Text;
using System.Text.RegularExpressions;

namespace NPOI.Util
{
    public class Character
    {
		// definition of a valid C# identifier: http://msdn.microsoft.com/en-us/library/aa664670(v=vs.71).aspx
		private const string FORMATTING_CHARACTER = @"\p{Cf}";
		private const string CONNECTING_CHARACTER = @"\p{Pc}";
		private const string DECIMAL_DIGIT_CHARACTER = @"\p{Nd}";
		private const string COMBINING_CHARACTER = @"\p{Mn}|\p{Mc}";
		private const string LETTER_CHARACTER = @"\p{Lu}|\p{Ll}|\p{Lt}|\p{Lm}|\p{Lo}|\p{Nl}";

		private const string IDENTIFIER_PART_CHARACTER = LETTER_CHARACTER + "|" +
														 DECIMAL_DIGIT_CHARACTER + "|" +
														 CONNECTING_CHARACTER + "|" +
														 COMBINING_CHARACTER + "|" +
														 FORMATTING_CHARACTER;

		private const string IDENTIFIER_PART_CHARACTERS = "(" + IDENTIFIER_PART_CHARACTER + ")+";
		private const string IDENTIFIER_START_CHARACTER = "(" + LETTER_CHARACTER + "|_)";

		private const string IDENTIFIER_OR_KEYWORD = IDENTIFIER_START_CHARACTER + "(" +
													 IDENTIFIER_PART_CHARACTERS + ")*";

		private static readonly Regex _validIdentifierRegex = new Regex("^" + IDENTIFIER_OR_KEYWORD + "$", RegexOptions.Compiled);

		private static readonly Regex _validIdentifierPartRegex = new Regex(Regex.Escape(IDENTIFIER_PART_CHARACTER), RegexOptions.Compiled);

		public static bool IsIdentifierPart(int src)
		{
			StringBuilder sb = new StringBuilder(char.ConvertFromUtf32(src));
			sb.Append(src);
			return _validIdentifierPartRegex.IsMatch(sb.ToString());
		}

		public static bool IsIdentifierPart(char src)
		{
			StringBuilder sb = new StringBuilder();
			sb.Append(src);
			return _validIdentifierPartRegex.IsMatch(sb.ToString());
		}

		public static int GetNumericValue(char src)
        {
            if (src >= '0' && src <= '9')
            {
                return (int)src;
            }
            if (src >= 'A' && src <= 'Z')
            {
                return ((int)src) - 55;
            }
            if (src >= 'a' && src <= 'z')
            {
                return ((int)src) - 87;
            }
            return -1;
        }

        //TODO: this should work but maybe not.
        public static bool isWhitespace(char src)
        {
            return char.IsWhiteSpace(src);
        }

		public static int CharCount(int codePoint)
		{
			return codePoint >= 0x10000 ? 2 : 1;
		}
	}
}
