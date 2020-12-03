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

        public virtual string Format(object obj, CultureInfo culture)
        {
            return obj.ToString();
        }
        public virtual string Format(object obj)
        {
            return Format(obj, new StringBuilder(), 0).ToString();
        }

        protected virtual StringBuilder Format(object obj, StringBuilder sb, int pos)
        {
            return sb.Append(obj);
        }
        public abstract StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture);

        //public abstract Object Parse(string source);
        public abstract object ParseObject(string source, int pos);
        public TimeZoneInfo TimeZone
        {
            get;
            set;
        }
    }

    /**
    * Format class for Excel's SSN Format. This class mimics Excel's built-in
    * SSN Formatting.
    *
    * @author James May
    */
    public class SSNFormat : FormatBase
    {
        public static readonly FormatBase Instance = new SSNFormat();
        private static string df = "000000000";
        private SSNFormat()
        {
            // enforce singleton
        }

        /** Format a number as an SSN */
        public override string Format(object obj, CultureInfo culture)
        {
            var result = ((double)obj).ToString(df, culture);
            var sb = new StringBuilder();
            sb.Append(result.Substring(0, 3)).Append('-');
            sb.Append(result.Substring(3, 2)).Append('-');
            sb.Append(result.Substring(5, 4));
            return sb.ToString();
        }

        protected override StringBuilder Format(object obj, StringBuilder toAppendTo, int pos)
        {
            return toAppendTo.Append(Format(obj, CultureInfo.CurrentCulture));
        }
        public override StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format((long)obj, culture));
        }

        public override object ParseObject(string source, int pos)
        {
            var tmp = source.Substring(pos);
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
        public static readonly FormatBase Instance = new ZipPlusFourFormat();
        private static string df = "000000000";
        private ZipPlusFourFormat()
        {
            // enforce singleton
        }

        /** Format a number as Zip + 4 */
        public override string Format(object obj, CultureInfo culture)
        {
            var result = ((double)obj).ToString(df, culture);
            return result.Substring(0, 5)+'-'+result.Substring(5, 4);
        }

        protected override StringBuilder Format(object obj, StringBuilder toAppendTo, int pos)
        {
            return toAppendTo.Append(Format(obj, CultureInfo.CurrentCulture));
        }

        public override StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format(obj, culture));
        }

        public override object ParseObject(string source, int pos)
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
        public static readonly FormatBase Instance = new PhoneFormat();
        private static string df = "##########";
        private PhoneFormat()
        {
            // enforce singleton
        }

        /** Format a number as a phone number */
        public override string Format(object obj, CultureInfo culture)
        {
            var result = ((double)obj).ToString(df, culture);
            var sb = new StringBuilder();
            String seg1, seg2, seg3;
            var len = result.Length;
            if (len <= 4)
            {
                return result;
            }

            seg3 = result.Substring(len - 4);
            var beginPos = Math.Max(0, len - 7);
            seg2 = result.Substring(Math.Max(0, len - 7), len - 4 - beginPos);
            beginPos = Math.Max(0, len - 10);
            seg1 = result.Substring(beginPos, Math.Max(0, len - 7) - beginPos);

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

        protected override StringBuilder Format(object obj, StringBuilder toAppendTo, int pos)
        {
            return toAppendTo.Append(Format(obj, CultureInfo.CurrentCulture));
        }
        public override StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format(obj, culture));
        }

        public override object ParseObject(string source, int pos)
        {
            var tmp = source.Substring(pos);
            return long.Parse(tmp, CultureInfo.InvariantCulture);
        }
    }

    public class DecimalFormat : FormatBase
    {
        public DecimalFormat()
        {

        }

        private string _pattern;
        private NumberFormatInfo _formatInfo;

        public DecimalFormat(string pattern)
        {
            if (pattern.IndexOf("'", StringComparison.Ordinal) != -1)
                throw new ArgumentException("invalid pattern");
            this._pattern = pattern;
        }
        public DecimalFormat(string pattern, NumberFormatInfo formatInfo)
            : this(pattern)
        {
            this._formatInfo = formatInfo;
        }
        public string Pattern
        {
            get { return _pattern; }
        }

        private static readonly Regex RegexFraction = new Regex("#+/#+");
        public override string Format(Object obj)
        {
            return Format(obj, CultureInfo.CurrentCulture);
        }
        public override string Format(object obj, CultureInfo culture)
        {
            //invalide fraction
            _pattern = RegexFraction.Replace(_pattern, "/");
            if (_formatInfo != null)
            {
                culture = (CultureInfo)culture.Clone();
                culture.NumberFormat = _formatInfo;
            }
                
            if (_pattern.IndexOf("'", StringComparison.Ordinal) != -1)
            {
                return Convert.ToDouble(obj, CultureInfo.InvariantCulture).ToString(culture);
            }
            else
            {
                var value = Convert.ToDouble(obj, CultureInfo.InvariantCulture);
                var ret = value.ToString(_pattern, culture);
                if (string.IsNullOrEmpty(ret))
                    ret = "0";
                return ret;
            }
        }

        public override StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format(obj, culture));
        }

        public override object ParseObject(string source, int pos)
        {
            return Decimal.Parse(source.Substring(pos), CultureInfo.CurrentCulture);
        }

        public bool ParseIntegerOnly {
            get { return false;}
        } 
    }

    public abstract class DateFormat : FormatBase
    {
        public const int FULL = 0;
        public const int LONG = 1;
        public const int MEDIUM = 2;
        public const int SHORT = 3;
        public const int DEFAULT = MEDIUM;

        public static string GetDateTimePattern(int dateStyle, int timeStyle, CultureInfo locale)
        {
            DateTimeFormatInfo dfi = locale.DateTimeFormat;
            string datePattern = GetDatePattern(dateStyle,locale);
            string timePattern = GetTimePattern(timeStyle, locale);

            if (locale.TextInfo.IsRightToLeft)
                return timePattern + " " + datePattern;
            else
                return datePattern + " " + timePattern;
        }
        public static string GetDatePattern(int dateStyle, CultureInfo locale)
        {
            DateTimeFormatInfo dfi = locale.DateTimeFormat;
            switch (dateStyle)
            {
                case DateFormat.SHORT:
                    return dfi.ShortDatePattern.Replace("yyyy", "yy").Replace("YYYY", "YY");
                case DateFormat.MEDIUM:
                    return dfi.ShortDatePattern;
                case DateFormat.LONG:
                    return dfi.LongDatePattern.Replace("dddd,", "").Trim();
                case DateFormat.FULL:
                    return dfi.LongDatePattern;
                default:
                    return dfi.ShortDatePattern;
            }
        }
        public static string GetTimePattern(int timeStyle, CultureInfo locale)
        {
            DateTimeFormatInfo dfi = locale.DateTimeFormat;
            switch (timeStyle)
            {
                case DateFormat.SHORT:
                    return dfi.ShortTimePattern;
                case DateFormat.MEDIUM:
                case DateFormat.LONG:
                case DateFormat.FULL:
                default:
                    return dfi.LongTimePattern;
            }
        }
    }
    public class SimpleDateFormat : DateFormat
    {
        private string _pattern;
        private DateTimeFormatInfo _formatData;
        private CultureInfo _culture;
        public SimpleDateFormat():this("", CultureInfo.CurrentCulture)
        {

        }

        public string Pattern
        {
            get { return _pattern; }
        }

        public SimpleDateFormat(string pattern, CultureInfo culture)
        {
            if (pattern == null || culture == null)
            {
                throw new ArgumentNullException();
            }
            this._pattern = pattern;
            this._formatData = (DateTimeFormatInfo)culture.DateTimeFormat.Clone();
            this._culture = culture;
        }
        
        public SimpleDateFormat(string pattern, DateTimeFormatInfo formatSymbols)
        {
            if (pattern == null || formatSymbols == null)
            {
                throw new ArgumentNullException();
            }
            this._pattern = pattern;
            this._formatData = (DateTimeFormatInfo)formatSymbols.Clone();
            this._culture = CultureInfo.CurrentCulture;
        }

        public SimpleDateFormat(string pattern)
        {
            this._pattern = pattern;
        }
        public override string Format(Object obj)
        {
            return Format(obj, CultureInfo.CurrentCulture);
        }
        public override string Format(object obj, CultureInfo culture)
        {
            DateTime dt = (DateTime)obj;
            if (TimeZone != null)
                dt = TimeZoneInfo.ConvertTime(dt, TimeZone);
            return  dt.ToString(_pattern, culture); 
        }

        public override StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(this.Format((DateTime)obj, culture));
        }

        public override object ParseObject(string source, int pos)
        {
            var dt = DateTime.Parse(source.Substring(pos), CultureInfo.InvariantCulture);
            return TimeZone != null ? TimeZoneInfo.ConvertTime(dt, TimeZone) : dt;
        }
        public DateTime Parse(string source)
        {
            var dt = DateTime.Parse(source, CultureInfo.InvariantCulture);
            return TimeZone != null ? TimeZoneInfo.ConvertTime(dt, TimeZone) : dt;
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
        private string str;
        public ConstantStringFormat(string s)
        {
            str = s;
        }
        public override string Format(object obj)
        {
            return str;
        }
        public override string Format(object obj, CultureInfo culture)
        {
            return str;
        }
        public override StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(str);
        }

        public override object ParseObject(string source, int pos)
        {
            return df.ParseObject(source, pos);
        }
    }
}
