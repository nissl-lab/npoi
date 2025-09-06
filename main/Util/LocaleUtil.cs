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
        private static TimeZoneInfo userTimeZone;

        [ThreadStatic]
        private static CultureInfo userLocale;

        /**
         * As time zone information is not stored in any format, it can be
         * set before any date calculations take place.
         * This setting is specific to the current thread.
         *
         * @param timezone the timezone under which date calculations take place
         */
        public static void SetUserTimeZone(TimeZoneInfo timezone)
        {
            userTimeZone = (timezone);
        }
        /**
         * @return the time zone which is used for date calculations, defaults to UTC
         */
        public static TimeZoneInfo GetUserTimeZoneInfo()
        {
            return userTimeZone ?? (userTimeZone = TimeZoneInfo.Local);
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
            return GetLocaleCalendar(GetUserTimeZoneInfo());
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
        public static DateTime GetLocaleCalendar(TimeZoneInfo timeZone)
        {
            return TimeZoneInfo.ConvertTime(DateTime.Now, timeZone);
        }

        /**
         * @return a calendar for the user locale and time zone
         */
        [Obsolete("The class TimeZone was marked obsolete, Use the Overload using TimeZoneInfo instead.")]
        public static DateTime GetLocaleCalendar(TimeZone timeZone)
        {
            return timeZone.ToLocalTime(DateTime.Now);
            //return Calendar.GetInstance(timeZone, GetUserLocale());
        }

        /// <summary>
        /// Decode the language ID from LCID value
        /// </summary>
        /// <param name="lcid">the LCID value</param>
        /// <returns>the locale/language ID</returns>
        public static string GetLocaleFromLCID(int lcid)
        {
            int languageId = lcid & 0xFFFF;
            switch (languageId) 
            {
                case 0x0001: return "ar";
                case 0x0002: return "bg";
                case 0x0003: return "ca";
                case 0x0004: return "zh-Hans";
                case 0x0005: return "cs";
                case 0x0006: return "da";
                case 0x0007: return "de";
                case 0x0008: return "el";
                case 0x0009: return "en";
                case 0x000a: return "es";
                case 0x000b: return "fi";
                case 0x000c: return "fr";
                case 0x000d: return "he";
                case 0x000e: return "hu";
                case 0x000f: return "is";
                case 0x0010: return "it";
                case 0x0011: return "ja";
                case 0x0012: return "ko";
                case 0x0013: return "nl";
                case 0x0014: return "no";
                case 0x0015: return "pl";
                case 0x0016: return "pt";
                case 0x0017: return "rm";
                case 0x0018: return "ro";
                case 0x0019: return "ru";
                case 0x001a: return "bs, hr, or sr";
                case 0x001b: return "sk";
                case 0x001c: return "sq";
                case 0x001d: return "sv";
                case 0x001e: return "th";
                case 0x001f: return "tr";
                case 0x0020: return "ur";
                case 0x0021: return "id";
                case 0x0022: return "uk";
                case 0x0023: return "be";
                case 0x0024: return "sl";
                case 0x0025: return "et";
                case 0x0026: return "lv";
                case 0x0027: return "lt";
                case 0x0028: return "tg";
                case 0x0029: return "fa";
                case 0x002a: return "vi";
                case 0x002b: return "hy";
                case 0x002c: return "az";
                case 0x002d: return "eu";
                case 0x002e: return "dsb or hsb";
                case 0x002f: return "mk";
                case 0x0030: return "st"; // reserved
                case 0x0031: return "ts"; // reserved
                case 0x0032: return "tn";
                case 0x0033: return "ve"; // reserved
                case 0x0034: return "xh";
                case 0x0035: return "zu";
                case 0x0036: return "af";
                case 0x0037: return "ka";
                case 0x0038: return "fo";
                case 0x0039: return "hi";
                case 0x003a: return "mt";
                case 0x003b: return "se";
                case 0x003c: return "ga";
                case 0x003d: return "yi"; // reserved
                case 0x003e: return "ms";
                case 0x003f: return "kk";
                case 0x0040: return "ky";
                case 0x0041: return "sw";
                case 0x0042: return "tk";
                case 0x0043: return "uz";
                case 0x0044: return "tt";
                case 0x0045: return "bn";
                case 0x0046: return "pa";
                case 0x0047: return "gu";
                case 0x0048: return "or";
                case 0x0049: return "ta";
                case 0x004a: return "te";
                case 0x004b: return "kn";
                case 0x004c: return "ml";
                case 0x004d: return "as";
                case 0x004e: return "mr";
                case 0x004f: return "sa";
                case 0x0050: return "mn";
                case 0x0051: return "bo";
                case 0x0052: return "cy";
                case 0x0053: return "km";
                case 0x0054: return "lo";
                case 0x0055: return "my"; // reserved
                case 0x0056: return "gl";
                case 0x0057: return "kok";
                case 0x0058: return "mni"; // reserved
                case 0x0059: return "sd";
                case 0x005a: return "syr";
                case 0x005b: return "si";
                case 0x005c: return "chr";
                case 0x005d: return "iu";
                case 0x005e: return "am";
                case 0x005f: return "tzm";
                case 0x0060: return "ks"; // reserved
                case 0x0061: return "ne";
                case 0x0062: return "fy";
                case 0x0063: return "ps";
                case 0x0064: return "fil";
                case 0x0065: return "dv";
                case 0x0066: return "bin"; // reserved
                case 0x0067: return "ff";
                case 0x0068: return "ha";
                case 0x0069: return "ibb"; // reserved
                case 0x006a: return "yo";
                case 0x006b: return "quz";
                case 0x006c: return "nso";
                case 0x006d: return "ba";
                case 0x006e: return "lb";
                case 0x006f: return "kl";
                case 0x0070: return "ig";
                case 0x0071: return "kr"; // reserved
                case 0x0072: return "om"; // reserved
                case 0x0073: return "ti";
                case 0x0074: return "gn"; // reserved
                case 0x0075: return "haw";
                case 0x0076: return "la"; // reserved
                case 0x0077: return "so"; // reserved
                case 0x0078: return "ii";
                case 0x0079: return "pap"; // reserved
                case 0x007a: return "arn";
                case 0x007b: return "invalid"; // Neither defined nor reserved
                case 0x007c: return "moh";
                case 0x007d: return "invalid"; // Neither defined nor reserved
                case 0x007e: return "br";
                case 0x007f: return "invalid"; // Reserved or invariant locale behavior
                case 0x0080: return "ug";
                case 0x0081: return "mi";
                case 0x0082: return "oc";
                case 0x0083: return "co";
                case 0x0084: return "gsw";
                case 0x0085: return "sah";
                case 0x0086: return "qut";
                case 0x0087: return "rw";
                case 0x0088: return "wo";
                case 0x0089: return "invalid"; // Neither defined nor reserved
                case 0x008a: return "invalid"; // Neither defined nor reserved
                case 0x008b: return "invalid"; // Neither defined nor reserved
                case 0x008c: return "prs";
                case 0x008d: return "invalid"; // Neither defined nor reserved
                case 0x008e: return "invalid"; // Neither defined nor reserved
                case 0x008f: return "invalid"; // Neither defined nor reserved
                case 0x0090: return "invalid"; // Neither defined nor reserved
                case 0x0091: return "gd";
                case 0x0092: return "ku";
                case 0x0093: return "quc"; // reserved
                case 0x0401: return "ar-SA";
                case 0x0402: return "bg-BG";
                case 0x0403: return "ca-ES";
                case 0x0404: return "zh-TW";
                case 0x0405: return "cs-CZ";
                case 0x0406: return "da-DK";
                case 0x0407: return "de-DE";
                case 0x0408: return "el-GR";
                case 0x0409: return "en-US";
                case 0x040a: return "es-ES_tradnl";
                case 0x040b: return "fi-FI";
                case 0x040c: return "fr-FR";
                case 0x040d: return "he-IL";
                case 0x040e: return "hu-HU";
                case 0x040f: return "is-IS";
                case 0x0410: return "it-IT";
                case 0x0411: return "ja-JP";
                case 0x0412: return "ko-KR";
                case 0x0413: return "nl-NL";
                case 0x0414: return "nb-NO";
                case 0x0415: return "pl-PL";
                case 0x0416: return "pt-BR";
                case 0x0417: return "rm-CH";
                case 0x0418: return "ro-RO";
                case 0x0419: return "ru-RU";
                case 0x041a: return "hr-HR";
                case 0x041b: return "sk-SK";
                case 0x041c: return "sq-AL";
                case 0x041d: return "sv-SE";
                case 0x041e: return "th-TH";
                case 0x041f: return "tr-TR";
                case 0x0420: return "ur-PK";
                case 0x0421: return "id-ID";
                case 0x0422: return "uk-UA";
                case 0x0423: return "be-BY";
                case 0x0424: return "sl-SI";
                case 0x0425: return "et-EE";
                case 0x0426: return "lv-LV";
                case 0x0427: return "lt-LT";
                case 0x0428: return "tg-Cyrl-TJ";
                case 0x0429: return "fa-IR";
                case 0x042a: return "vi-VN";
                case 0x042b: return "hy-AM";
                case 0x042c: return "az-Latn-AZ";
                case 0x042d: return "eu-ES";
                case 0x042e: return "hsb-DE";
                case 0x042f: return "mk-MK";
                case 0x0430: return "st-ZA"; // reserved
                case 0x0431: return "ts-ZA"; // reserved
                case 0x0432: return "tn-ZA";
                case 0x0433: return "ve-ZA"; // reserved
                case 0x0434: return "xh-ZA";
                case 0x0435: return "zu-ZA";
                case 0x0436: return "af-ZA";
                case 0x0437: return "ka-GE";
                case 0x0438: return "fo-FO";
                case 0x0439: return "hi-IN";
                case 0x043a: return "mt-MT";
                case 0x043b: return "se-NO";
                case 0x043d: return "yi-Hebr"; // reserved
                case 0x043e: return "ms-MY";
                case 0x043f: return "kk-KZ";
                case 0x0440: return "ky-KG";
                case 0x0441: return "sw-KE";
                case 0x0442: return "tk-TM";
                case 0x0443: return "uz-Latn-UZ";
                case 0x0444: return "tt-RU";
                case 0x0445: return "bn-IN";
                case 0x0446: return "pa-IN";
                case 0x0447: return "gu-IN";
                case 0x0448: return "or-IN";
                case 0x0449: return "ta-IN";
                case 0x044a: return "te-IN";
                case 0x044b: return "kn-IN";
                case 0x044c: return "ml-IN";
                case 0x044d: return "as-IN";
                case 0x044e: return "mr-IN";
                case 0x044f: return "sa-IN";
                case 0x0450: return "mn-MN";
                case 0x0451: return "bo-CN";
                case 0x0452: return "cy-GB";
                case 0x0453: return "km-KH";
                case 0x0454: return "lo-LA";
                case 0x0455: return "my-MM"; // reserved
                case 0x0456: return "gl-ES";
                case 0x0457: return "kok-IN";
                case 0x0458: return "mni-IN"; // reserved
                case 0x0459: return "sd-Deva-IN"; // reserved
                case 0x045a: return "syr-SY";
                case 0x045b: return "si-LK";
                case 0x045c: return "chr-Cher-US";
                case 0x045d: return "iu-Cans-CA";
                case 0x045e: return "am-ET";
                case 0x045f: return "tzm-Arab-MA"; // reserved
                case 0x0460: return "ks-Arab"; // reserved
                case 0x0461: return "ne-NP";
                case 0x0462: return "fy-NL";
                case 0x0463: return "ps-AF";
                case 0x0464: return "fil-PH";
                case 0x0465: return "dv-MV";
                case 0x0466: return "bin-NG"; // reserved
                case 0x0467: return "fuv-NG"; // reserved
                case 0x0468: return "ha-Latn-NG";
                case 0x0469: return "ibb-NG"; // reserved
                case 0x046a: return "yo-NG";
                case 0x046b: return "quz-BO";
                case 0x046c: return "nso-ZA";
                case 0x046d: return "ba-RU";
                case 0x046e: return "lb-LU";
                case 0x046f: return "kl-GL";
                case 0x0470: return "ig-NG";
                case 0x0471: return "kr-NG"; // reserved
                case 0x0472: return "om-Ethi-ET"; // reserved
                case 0x0473: return "ti-ET";
                case 0x0474: return "gn-PY"; // reserved
                case 0x0475: return "haw-US";
                case 0x0476: return "la-Latn"; // reserved
                case 0x0477: return "so-SO"; // reserved
                case 0x0478: return "ii-CN";
                case 0x0479: return "pap-x029"; // reserved
                case 0x047a: return "arn-CL";
                case 0x047c: return "moh-CA";
                case 0x047e: return "br-FR";
                case 0x0480: return "ug-CN";
                case 0x0481: return "mi-NZ";
                case 0x0482: return "oc-FR";
                case 0x0483: return "co-FR";
                case 0x0484: return "gsw-FR";
                case 0x0485: return "sah-RU";
                case 0x0486: return "qut-GT";
                case 0x0487: return "rw-RW";
                case 0x0488: return "wo-SN";
                case 0x048c: return "prs-AF";
                case 0x048d: return "plt-MG"; // reserved
                case 0x048e: return "zh-yue-HK"; // reserved
                case 0x048f: return "tdd-Tale-CN"; // reserved
                case 0x0490: return "khb-Talu-CN"; // reserved
                case 0x0491: return "gd-GB";
                case 0x0492: return "ku-Arab-IQ";
                case 0x0493: return "quc-CO"; // reserved
                case 0x0501: return "qps-ploc";
                case 0x05fe: return "qps-ploca";
                case 0x0801: return "ar-IQ";
                case 0x0803: return "ca-ES-valencia";
                case 0x0804: return "zh-CN";
                case 0x0807: return "de-CH";
                case 0x0809: return "en-GB";
                case 0x080a: return "es-MX";
                case 0x080c: return "fr-BE";
                case 0x0810: return "it-CH";
                case 0x0811: return "ja-Ploc-JP"; // reserved
                case 0x0813: return "nl-BE";
                case 0x0814: return "nn-NO";
                case 0x0816: return "pt-PT";
                case 0x0818: return "ro-MO"; // reserved
                case 0x0819: return "ru-MO"; // reserved
                case 0x081a: return "sr-Latn-CS";
                case 0x081d: return "sv-FI";
                case 0x0820: return "ur-IN"; // reserved
                case 0x0827: return "invalid"; // Neither defined nor reserved
                case 0x082c: return "az-Cyrl-AZ";
                case 0x082e: return "dsb-DE";
                case 0x0832: return "tn-BW";
                case 0x083b: return "se-SE";
                case 0x083c: return "ga-IE";
                case 0x083e: return "ms-BN";
                case 0x0843: return "uz-Cyrl-UZ";
                case 0x0845: return "bn-BD";
                case 0x0846: return "pa-Arab-PK";
                case 0x0849: return "ta-LK";
                case 0x0850: return "mn-Mong-CN";
                case 0x0851: return "bo-BT"; // reserved
                case 0x0859: return "sd-Arab-PK";
                case 0x085d: return "iu-Latn-CA";
                case 0x085f: return "tzm-Latn-DZ";
                case 0x0860: return "ks-Deva"; // reserved
                case 0x0861: return "ne-IN"; // reserved
                case 0x0867: return "ff-Latn-SN";
                case 0x086b: return "quz-EC";
                case 0x0873: return "ti-ER";
                case 0x09ff: return "qps-plocm";
                case 0x0c01: return "ar-EG";
                case 0x0c04: return "zh-HK";
                case 0x0c07: return "de-AT";
                case 0x0c09: return "en-AU";
                case 0x0c0a: return "es-ES";
                case 0x0c0c: return "fr-CA";
                case 0x0c1a: return "sr-Cyrl-CS";
                case 0x0c3b: return "se-FI";
                case 0x0c5f: return "tmz-MA"; // reserved
                case 0x0c6b: return "quz-PE";
                case 0x1001: return "ar-LY";
                case 0x1004: return "zh-SG";
                case 0x1007: return "de-LU";
                case 0x1009: return "en-CA";
                case 0x100a: return "es-GT";
                case 0x100c: return "fr-CH";
                case 0x101a: return "hr-BA";
                case 0x103b: return "smj-NO";
                case 0x1401: return "ar-DZ";
                case 0x1404: return "zh-MO";
                case 0x1407: return "de-LI";
                case 0x1409: return "en-NZ";
                case 0x140a: return "es-CR";
                case 0x140c: return "fr-LU";
                case 0x141a: return "bs-Latn-BA";
                case 0x143b: return "smj-SE";
                case 0x1801: return "ar-MA";
                case 0x1809: return "en-IE";
                case 0x180a: return "es-PA";
                case 0x180c: return "fr-MC";
                case 0x181a: return "sr-Latn-BA";
                case 0x183b: return "sma-NO";
                case 0x1c01: return "ar-TN";
                case 0x1c09: return "en-ZA";
                case 0x1c0a: return "es-DO";
                case 0x1c0c: return "invalid"; // Neither defined nor reserved
                case 0x1c1a: return "sr-Cyrl-BA";
                case 0x1c3b: return "sma-SE";
                case 0x2001: return "ar-OM";
                case 0x2008: return "invalid"; // Neither defined nor reserved
                case 0x2009: return "en-JM";
                case 0x200a: return "es-VE";
                case 0x200c: return "fr-RE"; // reserved
                case 0x201a: return "bs-Cyrl-BA";
                case 0x203b: return "sms-FI";
                case 0x2401: return "ar-YE";
                case 0x2409: return "en-029";
                case 0x240a: return "es-CO";
                case 0x240c: return "fr-CG"; // reserved
                case 0x241a: return "sr-Latn-RS";
                case 0x243b: return "smn-FI";
                case 0x2801: return "ar-SY";
                case 0x2809: return "en-BZ";
                case 0x280a: return "es-PE";
                case 0x280c: return "fr-SN"; // reserved
                case 0x281a: return "sr-Cyrl-RS";
                case 0x2c01: return "ar-JO";
                case 0x2c09: return "en-TT";
                case 0x2c0a: return "es-AR";
                case 0x2c0c: return "fr-CM"; // reserved
                case 0x2c1a: return "sr-Latn-ME";
                case 0x3001: return "ar-LB";
                case 0x3009: return "en-ZW";
                case 0x300a: return "es-EC";
                case 0x300c: return "fr-CI"; // reserved
                case 0x301a: return "sr-Cyrl-ME";
                case 0x3401: return "ar-KW";
                case 0x3409: return "en-PH";
                case 0x340a: return "es-CL";
                case 0x340c: return "fr-ML"; // reserved
                case 0x3801: return "ar-AE";
                case 0x3809: return "en-ID"; // reserved
                case 0x380a: return "es-UY";
                case 0x380c: return "fr-MA"; // reserved
                case 0x3c01: return "ar-BH";
                case 0x3c09: return "en-HK"; // reserved
                case 0x3c0a: return "es-PY";
                case 0x3c0c: return "fr-HT"; // reserved
                case 0x4001: return "ar-QA";
                case 0x4009: return "en-IN";
                case 0x400a: return "es-BO";
                case 0x4401: return "ar-Ploc-SA"; // reserved
                case 0x4409: return "en-MY";
                case 0x440a: return "es-SV";
                case 0x4801: return "ar-145"; // reserved
                case 0x4809: return "en-SG";
                case 0x480a: return "es-HN";
                case 0x4c09: return "en-AE"; // reserved
                case 0x4c0a: return "es-NI";
                case 0x5009: return "en-BH"; // reserved
                case 0x500a: return "es-PR";
                case 0x5409: return "en-EG"; // reserved
                case 0x540a: return "es-US";
                case 0x5809: return "en-JO"; // reserved
                case 0x5c09: return "en-KW"; // reserved
                case 0x6009: return "en-TR"; // reserved
                case 0x6409: return "en-YE"; // reserved
                case 0x641a: return "bs-Cyrl";
                case 0x681a: return "bs-Latn";
                case 0x6c1a: return "sr-Cyrl";
                case 0x701a: return "sr-Latn";
                case 0x703b: return "smn";
                case 0x742c: return "az-Cyrl";
                case 0x743b: return "sms";
                case 0x7804: return "zh";
                case 0x7814: return "nn";
                case 0x781a: return "bs";
                case 0x782c: return "az-Latn";
                case 0x783b: return "sma";
                case 0x7843: return "uz-Cyrl";
                case 0x7850: return "mn-Cyrl";
                case 0x785d: return "iu-Cans";
                case 0x7c04: return "zh-Hant";
                case 0x7c14: return "nb";
                case 0x7c1a: return "sr";
                case 0x7c28: return "tg-Cyrl";
                case 0x7c2e: return "dsb";
                case 0x7c3b: return "smj";
                case 0x7c43: return "uz-Latn";
                case 0x7c46: return "pa-Arab";
                case 0x7c50: return "mn-Mong";
                case 0x7c59: return "sd-Arab";
                case 0x7c5c: return "chr-Cher";
                case 0x7c5d: return "iu-Latn";
                case 0x7c5f: return "tzm-Latn";
                case 0x7c67: return "ff-Latn";
                case 0x7c68: return "ha-Latn";
                case 0x7c92: return "ku-Arab";
                default: return "invalid";
            }
        }
    }
}
