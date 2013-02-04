using System;
using System.Text;
using System.Globalization;
using System.Text.RegularExpressions;

namespace NPOI.SS.Util
{
    /// <summary>
    /// A substitute class for Format class in Java
    /// </summary>
    public abstract class FormatBase
    {
        public FormatBase()
        {

        }

        public virtual string Format(Object obj, CultureInfo culture)
        {
            return obj.ToString();
        }
        public abstract StringBuilder Format(Object obj, StringBuilder toAppendTo, CultureInfo culture);

        //public abstract Object Parse(string source);
        public abstract Object ParseObject(string source, int pos);

    }

    /**
    * Format class for Excel's SSN Format. This class mimics Excel's built-in
    * SSN Formatting.
    *
    * @author James May
    */
    public class SSNFormat : FormatBase
    {
        public static FormatBase instance = new SSNFormat();
        private static string df = "000000000";
        private SSNFormat()
        {
            // enforce singleton
        }

        /** Format a number as an SSN */
        public override String Format(object obj, CultureInfo culture)
        {
            String result = ((double)obj).ToString(df, culture);
            StringBuilder sb = new StringBuilder();
            sb.Append(result.Substring(0, 3)).Append('-');
            sb.Append(result.Substring(3, 2)).Append('-');
            sb.Append(result.Substring(5, 4));
            return sb.ToString();
        }

        public override StringBuilder Format(Object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format((long)obj, culture));
        }

        public override Object ParseObject(String source, int pos)
        {
            string tmp = source.Substring(pos);
            return long.Parse(tmp, CultureInfo.InvariantCulture);
        }
    }

    /**
     * Format class for Excel Zip + 4 Format. This class mimics Excel's
     * built-in Formatting for Zip + 4.
     * @author James May
     */
    public class ZipPlusFourFormat : FormatBase
    {
        public static FormatBase instance = new ZipPlusFourFormat();
        private static string df = "000000000";
        private ZipPlusFourFormat()
        {
            // enforce singleton
        }

        /** Format a number as Zip + 4 */
        public override String Format(object obj, CultureInfo culture)
        {
            String result = ((double)obj).ToString(df, culture);
            StringBuilder sb = new StringBuilder();
            sb.Append(result.Substring(0, 5)).Append('-');
            sb.Append(result.Substring(5, 4));
            return sb.ToString();
        }

        public override StringBuilder Format(Object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format(obj, culture));
        }

        public override Object ParseObject(String source, int pos)
        {
            string tmp = source.Substring(pos);
            return long.Parse(tmp, CultureInfo.InvariantCulture);
        }
    }

    /**
     * Format class for Excel phone number Format. This class mimics Excel's
     * built-in phone number Formatting.
     * @author James May
     */
    public class PhoneFormat : FormatBase
    {
        public static FormatBase instance = new PhoneFormat();
        private static string df = "##########";
        private PhoneFormat()
        {
            // enforce singleton
        }

        /** Format a number as a phone number */
        public override String Format(object obj, CultureInfo culture)
        {
            String result = ((double)obj).ToString(df, culture);
            StringBuilder sb = new StringBuilder();
            String seg1, seg2, seg3;
            int len = result.Length;
            if (len <= 4)
            {
                return result;
            }

            seg3 = result.Substring(len - 4);
            int beginpos = Math.Max(0, len - 7);
            seg2 = result.Substring(Math.Max(0, len - 7), len - 4 - beginpos);
            beginpos = Math.Max(0, len - 10);
            seg1 = result.Substring(beginpos, Math.Max(0, len - 7) - beginpos);

            if (seg1 != null && seg1.Trim().Length > 0)
            {
                sb.Append('(').Append(seg1).Append(") ");
            }
            if (seg2 != null && seg2.Trim().Length > 0)
            {
                sb.Append(seg2).Append('-');
            }
            sb.Append(seg3);
            return sb.ToString();
        }

        public override StringBuilder Format(Object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format(obj, culture));
        }

        public override Object ParseObject(String source, int pos)
        {
            string tmp = source.Substring(pos);
            return long.Parse(tmp, CultureInfo.InvariantCulture);
        }
    }

    public class DecimalFormat : FormatBase
    {
        public DecimalFormat()
        {

        }

        private string pattern;

        public DecimalFormat(string pattern)
        {
            if (pattern.IndexOf("'", StringComparison.Ordinal) != -1)
                throw new ArgumentException("invalid pattern");
            this.pattern = pattern;
        }
        public string Pattern
        {
            get
            {
                return pattern;
            }
        }

        private static readonly Regex RegexFraction = new Regex("#+/#+");
        public override string Format(object obj, CultureInfo culture)
        {
            //invalide fraction
            pattern = RegexFraction.Replace(pattern, "/");
            
            if (pattern.IndexOf("'", StringComparison.Ordinal) != -1)
            {
                //return ((double)obj).ToString();
                return Convert.ToDouble(obj, CultureInfo.InvariantCulture).ToString(culture);
            }
            else
            {
                return Convert.ToDouble(obj, CultureInfo.InvariantCulture).ToString(pattern, culture);
                //return ((double)obj).ToString(pattern) ;
            }
        }

        public override StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format((double)obj, culture));
        }

        public override object ParseObject(string source, int pos)
        {
            return System.Decimal.Parse(source.Substring(pos), CultureInfo.CurrentCulture);
        }
        private bool _parseIntegerOnly = false;
        public bool ParseIntegerOnly
        {
            get { return _parseIntegerOnly; }
            set { _parseIntegerOnly = value; }
        }

    }

    public class SimpleDateFormat : FormatBase
    {
        public SimpleDateFormat()
        {

        }

        protected string pattern;

        public SimpleDateFormat(string pattern)
        {
            this.pattern = pattern;
        }

        public override string Format(object obj, CultureInfo culture)
        {
            String result = ((DateTime)obj).ToString(pattern, culture); //DateTimeFormatInfo.InvariantInfo
            return result;
        }

        public override StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(this.Format((DateTime)obj, culture));
        }

        public override object ParseObject(string source, int pos)
        {
            return DateTime.Parse(source.Substring(pos), CultureInfo.InvariantCulture).ToUniversalTime();
        }
        public DateTime Parse(string source)
        {
            return DateTime.Parse(source, CultureInfo.InvariantCulture);
        }
    }
    /**
     * Format class that handles Excel style fractions, such as "# #/#" and "#/###"
     */
    public class FractionFormat : FormatBase
    {
        private String formatstr = string.Empty;
        public FractionFormat(String s)
        {
            formatstr = s;
        }
        private static Regex regex = new Regex("(#+)([^#]*?)(#+/#+)([^#]*?)");
        public String Format(double num)
        {

            double doubleValue = num;

            // Format may be p or p;n or p;n;z (okay we never get a z).
            // Fall back to p when n or z is not specified.
            String[] formatBits = formatstr.Split(";".ToCharArray());
            int f = doubleValue > 0.0 ? 0 : doubleValue < 0.0 ? 1 : 2;
            String str = (f < formatBits.Length) ? formatBits[f] : formatBits[0];

            double wholePart = Math.Floor(Math.Abs(doubleValue));
            double decPart = Math.Abs(doubleValue) - wholePart;
            if (wholePart + decPart == 0)
            {
                return "0";
            }
            if (doubleValue < 0.0)
            {
                wholePart *= -1.0;
            }
            //str = regex.Replace(str, ReplaceWS);
            // Split the format string into decimal and fraction parts
            String[] parts = str.Replace("  *", " ").Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            String[] fractParts;
            if (parts.Length == 2)
            {
                fractParts = parts[1].Split("/".ToCharArray());
            }
            else
            {
                fractParts = str.Split("/".ToCharArray());
            }

            // Excel supports both #/# and ?/?, but Java only the former
            for (int i = 0; i < fractParts.Length; i++)
            {
                fractParts[i] = fractParts[i].Replace('?', '#');
            }

            if (fractParts.Length == 2)
            {
                int fractPart1Length = Math.Min(CountHashes(fractParts[1]), 4); // Any more than 3 and we go around the loops for ever
                double minVal = 1.0;
                double currDenom = Math.Pow(10, fractPart1Length) - 1d;
                double currNeum = 0;
                for (int i = (int)(Math.Pow(10, fractPart1Length) - 1d); i > 0; i--)
                {
                    for (int i2 = (int)(Math.Pow(10, fractPart1Length) - 1d); i2 > 0; i2--)
                    {
                        if (minVal >= Math.Abs((double)i2 / (double)i - decPart))
                        {
                            currDenom = i;
                            currNeum = i2;
                            minVal = Math.Abs((double)i2 / (double)i - decPart);
                        }
                    }
                }
                DecimalFormat neumFormatter = new DecimalFormat(fractParts[0]);
                DecimalFormat denomFormatter = new DecimalFormat(fractParts[1]);
                CultureInfo cul = CultureInfo.CurrentCulture;
                if (parts.Length == 2)
                {
                    DecimalFormat wholeFormatter = new DecimalFormat(parts[0]);
                    String result = wholeFormatter.Format(wholePart, cul) + " " + neumFormatter.Format(currNeum, cul) + "/" + denomFormatter.Format(currDenom, cul);
                    return result;
                }
                else
                {
                    String result = neumFormatter.Format(currNeum + (currDenom * wholePart), cul) + "/" + denomFormatter.Format(currDenom, cul);
                    return result;
                }
            }
            else
            {
                throw new ArgumentException("Fraction must have 2 parts, found " + fractParts.Length + " for fraction format " + this.formatstr);
            }
        }

        private string ReplaceWS(Match m)
        // Replace each Regex cc match with the number of the occurrence.
        {
            return m.Groups[0].Value + " " + m.Groups[2].Value + m.Groups[3].Value;
        }


        private int CountHashes(String format)
        {
            int count = 0;
            for (int i = format.Length - 1; i >= 0; i--)
            {
                if (format[(i)] == '#')
                {
                    count++;
                }
            }
            return count;
        }
        public override string Format(object obj, CultureInfo culture)
        {
            return Format((double)obj);
        }
        public override StringBuilder Format(Object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format((double)obj));
        }

        public override Object ParseObject(String source, int pos)
        {
            throw new NotImplementedException("Reverse parsing not supported");
        }
    }
    /**
     * Format class that does nothing and always returns a constant string.
     *
     * This format is used to simulate Excel's handling of a format string
     * of all # when the value is 0. Excel will output "", Java will output "0".
     *
     * @see DataFormatter#createFormat(double, int, String)
     */
    public class ConstantStringFormat : FormatBase
    {
        private static DecimalFormat df = new DecimalFormat("##########");
        private String str;
        public ConstantStringFormat(String s)
        {
            str = s;
        }
        public override string Format(object obj, CultureInfo culture)
        {
            return str;
        }
        public override StringBuilder Format(Object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(str);
        }

        public override Object ParseObject(String source, int pos)
        {
            return df.ParseObject(source, pos);
        }
    }
}
