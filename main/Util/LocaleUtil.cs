using System;
using System.Globalization;

namespace NPOI.Util
{
    /**
 * This utility class is used to set locale and time zone settings beside
 * of the JDK internal {@link java.util.Locale#setDefault(Locale)} and
 * {@link java.util.TimeZone#setDefault(TimeZone)} methods, because
 * the locale/time zone specific handling of certain office documents -
 * maybe for different time zones / locales ... - shouldn't affect
 * other java components.
 * 
 * The settings are saved in a {@link java.lang.ThreadLocal},
 * so they only apply to the current thread and can't be set globally.
 */
    public class LocaleUtil
    {
        /**
         * Excel doesn't store TimeZone information in the file, so if in doubt,
         *  use UTC to perform calculations
         */
        public static TimeZoneInfo TIMEZONE_UTC = TimeZoneInfo.Utc;

        /**
         * Default encoding for unknown byte encodings of native files
         * (at least it's better than to rely on a platform dependent encoding
         * for legacy stuff ...)
         */
        public static string CHARSET_1252 = CodePageUtil.CodepageToEncoding(CodePageUtil.CP_WINDOWS_1252);// ("CP1252");

        [ThreadStatic]
        private static TimeZone userTimeZone;

        [ThreadStatic]
        private static CultureInfo userLocale;
        //private static ThreadLocal<Locale> userLocale = new ThreadLocal<Locale>();

        /**
         * As time zone information is not stored in any format, it can be
         * set before any date calculations take place.
         * This setting is specific to the current thread.
         *
         * @param timezone the timezone under which date calculations take place
         */
        public static void SetUserTimeZone(TimeZone timezone)
        {
            userTimeZone = (timezone);
        }

        /**
         * @return the time zone which is used for date calculations, defaults to UTC
         */
        public static TimeZone GetUserTimeZone()
        {
            return userTimeZone ?? (userTimeZone = TimeZone.CurrentTimeZone);
        }

        /**
         * Sets default user locale.
         * This setting is specific to the current thread.
         */
        public static void SetUserLocale(CultureInfo locale)
        {
            userLocale = (locale);
        }

        /**
         * @return the default user locale, defaults to {@link Locale#ROOT}
         */
        public static CultureInfo GetUserLocale()
        {
            return userLocale ?? (userLocale = CultureInfo.CurrentCulture);
        }

        /**
         * @return a calendar for the user locale and time zone
         */
        public static DateTime GetLocaleCalendar()
        {
            return GetLocaleCalendar(GetUserTimeZone());
        }

        /**
         * Convenience method - month is 0-based as in java.util.Calendar
         *
         * @param year
         * @param month
         * @param day
         * @return a calendar for the user locale and time zone, and the given date
         */
        public static DateTime GetLocaleCalendar(int year, int month, int day)
        {
            return GetLocaleCalendar(year, month, day, 0, 0, 0);
        }

        /**
         * Convenience method - month is 0-based as in java.util.Calendar
         *
         * @param year
         * @param month
         * @param day
         * @param hour
         * @param minute
         * @param second
         * @return a calendar for the user locale and time zone, and the given date
         */
        public static DateTime GetLocaleCalendar(int year, int month, int day, int hour, int minute, int second)
        {
            //DateTime cal = GetLocaleCalendar();
            //cal.Set(year, month, day, hour, minute, second);
            //cal.Clear(Calendar.MILLISECOND);
            if (month < 0 || day < 0)
            {
                throw new ArgumentOutOfRangeException();
            }
            DateTime cal;
            if (day == 0)
            {
                cal = new DateTime(year, month, 1, hour, minute, second);
                cal = cal.AddDays(-1);
            }
            else
            {
                cal = new DateTime(year, month, day, hour, minute, second);
            }
            return cal;
        }

        /**
         * @return a calendar for the user locale and time zone
         */
        public static DateTime GetLocaleCalendar(TimeZone timeZone)
        {
            return timeZone.ToLocalTime(DateTime.Now);
            //return Calendar.GetInstance(timeZone, GetUserLocale());
        }
    }
}
