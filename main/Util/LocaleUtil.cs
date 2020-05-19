using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;

namespace NPOI.Util
{
    public class LocaleUtil
    {
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
