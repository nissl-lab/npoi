using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCases.SS.UserModel
{
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using System.Globalization;
    using NPOI.SS.Util;

    [TestFixture]
    public class TestExcelStyleDateFormatter
    {
        private static String EXCEL_DATE_FORMAT = "MMMMM";
        private static CultureInfo ROOT = CultureInfo.GetCultureInfo("en-US");
        /// <summary>
        /// [Bug 60369] Month format 'MMMMM' issue with TEXT-formula and Java 8
        /// </summary>
        [Test]
        public void Test60369()
        {

            // Setting up the locale to be tested together with a list of asserted unicode-formatted results and Put them in a map.
            //Locale germanLocale = Locale.GERMAN;
            CultureInfo germanLocale = CultureInfo.GetCultureInfo("de-DE");

            //Locale russianLocale = new Locale("ru", "RU");
            CultureInfo russianLocale = CultureInfo.GetCultureInfo("ru-RU");

            //Locale austrianLocale = new Locale("de", "AT");
            CultureInfo austrianLocale = CultureInfo.GetCultureInfo("de-AT");

            //Locale englishLocale = Locale.UK;
            CultureInfo englishLocale = CultureInfo.GetCultureInfo("en-GB");

            //Locale frenchLocale = Locale.FRENCH;
            CultureInfo frenchLocale = CultureInfo.GetCultureInfo("fr-FR");

            //Locale chineseLocale = Locale.CHINESE;
            CultureInfo chineseLocale = CultureInfo.GetCultureInfo("zh-Hans");

            //Locale turkishLocale = new Locale("tr", "TR");
            CultureInfo turkishLocale = CultureInfo.GetCultureInfo("tr-TR");

            //Locale hungarianLocale = new Locale("hu", "HU");
            CultureInfo hungarianLocale = CultureInfo.GetCultureInfo("hu-HU");

            //Locale indianLocale = new Locale("en", "IN");
            CultureInfo indianLocale = CultureInfo.GetCultureInfo("en-IN");

            //Locale indonesianLocale = new Locale("in", "ID");
            CultureInfo indonesianLocale = CultureInfo.GetCultureInfo("id-ID");

            Dictionary<CultureInfo, String> testMap = new()
            {
                { chineseLocale, "123456789111"}, // "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u5341\u5341"="一二三四五六七八九十十十" ，"123456789111"
                { germanLocale, "JFMAMJJASOND" },
                { russianLocale, "\u044f\u0444\u043c\u0430\u043c\u0438\u0438\u0430\u0441\u043e\u043d\u0434" },
                { austrianLocale, "JFMAMJJASOND" },
                { englishLocale, "JFMAMJJASOND" },
                { frenchLocale, "jfmamjjasond" },
                { turkishLocale, "\u004f\u015e\u004d\u004e\u004d\u0048\u0054\u0041\u0045\u0045\u004b\u0041" },
                { hungarianLocale, "\u006a\u0066\u006d\u00e1\u006d\u006a\u006a\u0061\u0073\u006f\u006e\u0064" },
                { indianLocale, "JFMAMJJASOND" },
                { indonesianLocale, "JFMAMJJASOND" }
            };
            String[] dates = {
            "1980-01-12", "1995-02-11", "2045-03-10", "2016-04-09", "2017-05-08",
            "1945-06-07", "1998-07-06", "2099-08-05", "1988-09-04", "2023-10-03", "1978-11-02", "1890-12-01"
        };
            // We have to Set up dates as well.
            List<DateTime> testDates = new();
            Array.ForEach(dates, (s) => testDates.Add(DateTime.Parse(s)));
            

            // Let's iterate over the test Setup.
            foreach (CultureInfo locale in testMap.Keys)
            {
                //System.err.Println("Locale: " + locale);
                ExcelStyleDateFormatter formatter = new ExcelStyleDateFormatter(EXCEL_DATE_FORMAT, locale.DateTimeFormat);
                for (int i = 0; i < 12; i++)
                {
                    // Call the method to be tested!
                    String result =
                            formatter.Format(testDates[i],
                                    new StringBuilder(),
                                    locale).ToString();
                                    //new FieldPosition(java.text.DateFormat.MONTH_FIELD)).ToString();
                    //System.err.Println(result +  " - " + GetUnicode(result[0]));
                    Assert.AreEqual(GetUnicode(testMap[locale][i]), GetUnicode(result[0]),
                        "current culture:"+ locale.ToString() + ", date="+ testDates[i]);
                }
            }
        }

        private String GetUnicode(char c)
        {
            return "\\u" + HexDump.ToHex(c | 0x10000).Substring(1);
        }

        [Test]
        public void TestConstruct()
        {
            Assert.IsNotNull(new ExcelStyleDateFormatter(EXCEL_DATE_FORMAT, LocaleUtil.GetUserLocale()));
            Assert.IsNotNull(new ExcelStyleDateFormatter(EXCEL_DATE_FORMAT));
        }

        [Test]
        public void TestWithLocale()
        {

            CultureInfo before = LocaleUtil.GetUserLocale();
            try
            {
                LocaleUtil.SetUserLocale(CultureInfo.GetCultureInfo("de-DE"));
                String dateStr = new ExcelStyleDateFormatter(EXCEL_DATE_FORMAT).Format(
                        new SimpleDateFormat("yyyy-MM-dd", ROOT).Parse("2016-03-26"),
                        CultureInfo.GetCultureInfo("de-DE"));
                Assert.AreEqual("M", dateStr);
            }
            finally
            {
                LocaleUtil.SetUserLocale(before);
            }
        }

        [Test]
        public void TestWithPattern()
        {
            String dateStr = new ExcelStyleDateFormatter("yyyy|" + EXCEL_DATE_FORMAT + "|").Format(
                    new SimpleDateFormat("yyyy-MM-dd", ROOT).Parse("2016-03-26"), ROOT);
            Assert.AreEqual("2016|M|", dateStr);
        }
    }
}