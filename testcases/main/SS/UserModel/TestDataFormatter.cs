/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.SS.UserModel
{
    using System;
    using NUnit.Framework;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using System.Globalization;
    using NPOI.SS.Format;
    using TestCases.HSSF;

    /**
     * Tests of {@link DataFormatter}
     *
     * See {@link TestHSSFDataFormatter} too for
     *  more Tests.
     */
    [TestFixture]
    public class TestDataFormatter
    {
        private static double _15_MINUTES = 0.041666667;
        [SetUp]
        public void SetUpClass()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            // some pre-checks to hunt for a problem in the Maven build
            // these checks ensure that the correct locale is set, so a Assert.Failure here
            // usually indicates an invalid locale during test-execution
            Assert.IsFalse(DateUtil.IsADateFormat(-1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            //Locale ul = LocaleUtil.getUserLocale();
            //assertTrue(Locale.ROOT.equals(ul) || Locale.getDefault().equals(ul));
            String textValue = NumberToTextConverter.ToText(1234.56);
            Assert.AreEqual(-1, textValue.IndexOf('E'));
            Object cellValueO = 1234.56d;
            /*CellFormat cellFormat = new CellFormat("_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-");
            CellFormatResult result = cellFormat.apply(cellValueO);
            Assert.AreEqual("    1,234.56 ", result.text);*/
            CellFormat cfmt = CellFormat.GetInstance("_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-");
            CellFormatResult result = cfmt.Apply(cellValueO);
            Assert.AreEqual("    1,234.56 ", result.Text,
                "This Assert.Failure can indicate that the wrong locale is used during test-execution, ensure you run with english/US via -Duser.language=en -Duser.country=US");
        }
        /**
         * Test that we use the specified locale when deciding
         *   how to format normal numbers
         */
        [Test]
        public void TestLocale()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
            DataFormatter dfFR = new DataFormatter(CultureInfo.GetCultureInfo("fr-FR"));

            Assert.AreEqual("1234", dfUS.FormatRawCellContents(1234, -1, "@"));
            Assert.AreEqual("1234", dfFR.FormatRawCellContents(1234, -1, "@"));
            
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.34, -1, "@"));
            Assert.AreEqual("12,34", dfFR.FormatRawCellContents(12.34, -1, "@"));
        }
        /**
         * At the moment, we don't decode the locale strings into
         *  a specific locale, but we should format things as if
         *  the locale (eg '[$-1010409]') isn't there
         */
        [Test]
        public void TestLocaleBasedFormats()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));

            // Standard formats
            Assert.AreEqual("63", dfUS.FormatRawCellContents(63.0, -1, "[$-1010409]General"));
            Assert.AreEqual("63", dfUS.FormatRawCellContents(63.0, -1, "[$-1010409]@"));

            // Regular numeric style formats
            Assert.AreEqual("63", dfUS.FormatRawCellContents(63.0, -1, "[$-1010409]##"));
            Assert.AreEqual("63", dfUS.FormatRawCellContents(63.0, -1, "[$-1010409]00"));

        }
        /**
         * Test that we use the specified locale when deciding
         *   how to format normal numbers
         */
        [Test]
        public void TestGrouping()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
            DataFormatter dfDE = new DataFormatter(CultureInfo.GetCultureInfo("de-DE"));

            Assert.AreEqual("1,234.57", dfUS.FormatRawCellContents(1234.567, -1, "#,##0.00"));
            Assert.AreEqual("1'234.57", dfUS.FormatRawCellContents(1234.567, -1, "#'##0.00"));
            Assert.AreEqual("1 234.57", dfUS.FormatRawCellContents(1234.567, -1, "# ##0.00"));

            Assert.AreEqual("1.234,57", dfDE.FormatRawCellContents(1234.567, -1, "#,##0.00"));
            Assert.AreEqual("1'234,57", dfDE.FormatRawCellContents(1234.567, -1, "#'##0.00"));
            Assert.AreEqual("1 234,57", dfDE.FormatRawCellContents(1234.567, -1, "# ##0.00"));
        }

        /**
         * Ensure that colours Get correctly
         *  zapped from within the format strings
         */
        [Test]
        public void TestColours()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));

            String[] formats = new String[] {
             "##.##",
             "[WHITE]##.##",
             "[BLACK]##.##;[RED]-##.##",
             "[COLOR11]##.##;[COLOR 43]-##.00",
       };
            foreach (String format in formats)
            {
                Assert.AreEqual(
                      "12.34",
                      dfUS.FormatRawCellContents(12.343, -1, format),
                      "Wrong format for: " + format
                );
                Assert.AreEqual(
                      "-12.34",
                      dfUS.FormatRawCellContents(-12.343, -1, format),
                                            "Wrong format for: " + format

                );
            }

            // Ensure that random square brackets remain
            Assert.AreEqual("12.34[a]", dfUS.FormatRawCellContents(12.343, -1, "##.##[a]"));
            Assert.AreEqual("[ab]12.34[x]", dfUS.FormatRawCellContents(12.343, -1, "[ab]##.##[x]"));
        }
        [Test]
        public void TestColoursAndBrackets()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));

            // Without currency symbols
            String[] formats = new String[] {
             "#,##0.00;[Blue](#,##0.00)",
       };
            foreach (String format in formats)
            {
                Assert.AreEqual(
                      
                      "12.34",
                      dfUS.FormatRawCellContents(12.343, -1, format),
                      "Wrong format for: " + format
                );
                Assert.AreEqual(
                      
                      "(12.34)",
                      dfUS.FormatRawCellContents(-12.343, -1, format),
                      "Wrong format for: " + format
                );
            }

            // With
            formats = new String[] {
             "$#,##0.00;[Red]($#,##0.00)"
       };
            foreach (String format in formats)
            {
                Assert.AreEqual(
                      "$12.34",
                      dfUS.FormatRawCellContents(12.343, -1, format),
                      "Wrong format for: " + format
                );
                Assert.AreEqual(
                      "($12.34)",
                      dfUS.FormatRawCellContents(-12.343, -1, format),
                      "Wrong format for: " + format
                );
            }
        }

        [Test]
        public void TestConditionalRanges()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
        
            String format = "[>=10]#,##0;[<10]0.0";
            Assert.AreEqual("17,876", dfUS.FormatRawCellContents(17876.000, -1, format), "Wrong format for " + format);
            Assert.AreEqual("9.7", dfUS.FormatRawCellContents(9.71, -1, format), "Wrong format for " + format);
        }

        /**
         * Test how we handle negative and zeros.
         * Note - some Tests are disabled as DecimalFormat
         *  and Excel differ, and workarounds are not
         *  yet in place for all of these
         */
        [Test]
        public void TestNegativeZero()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));

            String all2dp = "00.00";
            String alln1dp = "(00.0)";
            String p1dp_n1dp = "00.0;(00.0)";
            String p2dp_n1dp = "00.00;(00.0)";
            String p2dp_n1dp_z0 = "00.00;(00.0);0";
            String all2dpTSP = "00.00_x";
            String p2dp_n2dpTSP = "00.00_x;(00.00)_x";

            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, all2dp));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, p2dp_n1dp));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, p2dp_n1dp_z0));

            Assert.AreEqual("(12.3)", dfUS.FormatRawCellContents(12.343, -1, alln1dp));
            Assert.AreEqual("-(12.3)", dfUS.FormatRawCellContents(-12.343, -1, alln1dp));
            Assert.AreEqual("12.3", dfUS.FormatRawCellContents(12.343, -1, p1dp_n1dp));
            Assert.AreEqual("(12.3)", dfUS.FormatRawCellContents(-12.343, -1, p1dp_n1dp));

            Assert.AreEqual("-12.34", dfUS.FormatRawCellContents(-12.343, -1, all2dp));
            // TODO - fix case of negative subpattern differing from the
            //  positive one by more than just the prefix+suffix, which
            //  is all DecimalFormat supports...
            //       Assert.AreEqual("(12.3)", dfUS.FormatRawCellContents(-12.343, -1, p2dp_n1dp));
            //       Assert.AreEqual("(12.3)", dfUS.FormatRawCellContents(-12.343, -1, p2dp_n1dp_z0));

            Assert.AreEqual("00.00", dfUS.FormatRawCellContents(0, -1, all2dp));
            Assert.AreEqual("00.00", dfUS.FormatRawCellContents(0, -1, p2dp_n1dp));
            Assert.AreEqual("0", dfUS.FormatRawCellContents(0, -1, p2dp_n1dp_z0));

            // Spaces are skipped
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, all2dpTSP));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, p2dp_n2dpTSP));
            Assert.AreEqual("(12.34)", dfUS.FormatRawCellContents(-12.343, -1, p2dp_n2dpTSP));
            //String p2dp_n1dpTSP = "00.00_x;(00.0)_x";
            //       Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, p2dp_n1dpTSP));
            //       Assert.AreEqual("(12.3)", dfUS.FormatRawCellContents(-12.343, -1, p2dp_n1dpTSP));
        }

        /**
         * Test that we correctly handle fractions in the
         *  format string, eg # #/#
         */
        [Test]
        public void TestFractions()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));

            // Excel often prefers "# #/#"
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "# #/#"));
            Assert.AreEqual("321 26/81", dfUS.FormatRawCellContents(321.321, -1, "# #/##"));
            Assert.AreEqual("26027/81", dfUS.FormatRawCellContents(321.321, -1, "#/##"));

            // OOo seems to like the "# ?/?" form
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "# ?/?"));
            Assert.AreEqual("321 26/81", dfUS.FormatRawCellContents(321.321, -1, "# ?/??"));
            Assert.AreEqual("26027/81", dfUS.FormatRawCellContents(321.321, -1, "?/??"));

            // p;n;z;s parts
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "# #/#;# ##/#;0;xxx"));
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(-321.321, -1, "# #/#;# ##/#;0;xxx")); // Note the lack of - sign!
            Assert.AreEqual("0", dfUS.FormatRawCellContents(0, -1, "# #/#;# ##/#;0;xxx"));
            //     Assert.AreEqual(".",        dfUS.FormatRawCellContents(0,       -1, "# #/#;# ##/#;#.#;xxx")); // Currently shows as 0. not .

            // Custom formats with text are not currently supported
            Assert.AreEqual("+ve", dfUS.FormatRawCellContents(1, -1, "+ve;-ve;zero;xxx"));
            Assert.AreEqual("-ve", dfUS.FormatRawCellContents(-1, -1, "-ve;-ve;zero;xxx"));
            Assert.AreEqual("zero", dfUS.FormatRawCellContents(0, -1, "zero;-ve;zero;xxx"));

            // Custom formats - check text is stripped, including multiple spaces
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "#   #/#"));
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "#\"  \" #/#"));
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "#\"FRED\" #/#"));
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "#\\ #/#"));
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "# \\q#/#"));

            // Cases that were very slow
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "0\" \"?/?;?/?")); // 0" "?/?;?/?     - length of -ve part was used
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "0 \"#\"\\#\\#?/?")); // 0 "#"\#\#?/? - length of text was used

            Assert.AreEqual("321 295/919", dfUS.FormatRawCellContents(321.321, -1, "# #/###"));
            Assert.AreEqual("321 321/1000", dfUS.FormatRawCellContents(321.321, -1, "# #/####")); // Code limits to #### as that is as slow as we want to get
            Assert.AreEqual("321 321/1000", dfUS.FormatRawCellContents(321.321, -1, "# #/##########"));

            // Not a valid fraction formats (too many #/# or ?/?) - hence the strange expected results
            /*Assert.AreEqual("321 / ?/?", dfUS.FormatRawCellContents(321.321, -1, "# #/# ?/?"));
            Assert.AreEqual("321 / /", dfUS.FormatRawCellContents(321.321, -1, "# #/# #/#"));
            Assert.AreEqual("321 ?/? ?/?", dfUS.FormatRawCellContents(321.321, -1, "# ?/? ?/?"));
            Assert.AreEqual("321 ?/? / /", dfUS.FormatRawCellContents(321.321, -1, "# ?/? #/# #/#"));
            */

            //Bug54686 patch sets default behavior of # #/## if there is a failure to parse
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "# #/# ?/?"));
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "# #/# #/#"));
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "# ?/? ?/?"));
            Assert.AreEqual("321 1/3", dfUS.FormatRawCellContents(321.321, -1, "# ?/? #/# #/#"));

            // Where +ve has a fraction, but -ve doesnt, we currently show both
            Assert.AreEqual("123 1/3", dfUS.FormatRawCellContents(123.321, -1, "0 ?/?;0"));
            //Assert.AreEqual("123", dfUS.FormatRawCellContents(-123.321, -1, "0 ?/?;0"));

            //Bug54868 patch has a hit on the first string before the ";"
            Assert.AreEqual("-123 1/3", dfUS.FormatRawCellContents(-123.321, -1, "0 ?/?;0"));
            Assert.AreEqual("123 1/3", dfUS.FormatRawCellContents(123.321, -1, "0 ?/?;0"));

            //Bug53150 formatting a whole number with fractions should just give the number
            Assert.AreEqual("1", dfUS.FormatRawCellContents(1.0, -1, "# #/#"));
            Assert.AreEqual("11", dfUS.FormatRawCellContents(11.0, -1, "# #/#"));
        }

        /**
         * Test that _x (blank with the space taken by "x")
         *  and *x (fill to the column width with "x"s) are
         *  correctly ignored by us.
         */
        [Test]
        public void TestPaddingSpaces()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "##.##_ "));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "##.##_1"));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "##.##_)"));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "_-##.##"));

            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "##.##* "));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "##.##*1"));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "##.##*)"));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "*-##.##"));
        }

        /**
         * DataFormatter is the CSV mode preserves spaces
         */
        [Test]
        public void TestPaddingSpacesCSV()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"), true);
            Assert.AreEqual("12.34 ", dfUS.FormatRawCellContents(12.343, -1, "##.##_ "));
            Assert.AreEqual("-12.34 ", dfUS.FormatRawCellContents(-12.343, -1, "##.##_ "));
            Assert.AreEqual(". ", dfUS.FormatRawCellContents(0.0, -1, "##.##_ "));
            Assert.AreEqual("12.34 ", dfUS.FormatRawCellContents(12.343, -1, "##.##_1"));
            Assert.AreEqual("-12.34 ", dfUS.FormatRawCellContents(-12.343, -1, "##.##_1"));
            Assert.AreEqual(". ", dfUS.FormatRawCellContents(0.0, -1, "##.##_1"));
            Assert.AreEqual("12.34 ", dfUS.FormatRawCellContents(12.343, -1, "##.##_)"));
            Assert.AreEqual("-12.34 ", dfUS.FormatRawCellContents(-12.343, -1, "##.##_)"));
            Assert.AreEqual(". ", dfUS.FormatRawCellContents(0.0, -1, "##.##_)"));
            Assert.AreEqual(" 12.34", dfUS.FormatRawCellContents(12.343, -1, "_-##.##"));
            Assert.AreEqual("- 12.34", dfUS.FormatRawCellContents(-12.343, -1, "_-##.##"));
            Assert.AreEqual(" .", dfUS.FormatRawCellContents(0.0, -1, "_-##.##"));

            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "##.##* "));
            Assert.AreEqual("-12.34", dfUS.FormatRawCellContents(-12.343, -1, "##.##* "));
            Assert.AreEqual(".", dfUS.FormatRawCellContents(0.0, -1, "##.##* "));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "##.##*1"));
            Assert.AreEqual("-12.34", dfUS.FormatRawCellContents(-12.343, -1, "##.##*1"));
            Assert.AreEqual(".", dfUS.FormatRawCellContents(0.0, -1, "##.##*1"));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "##.##*)"));
            Assert.AreEqual("-12.34", dfUS.FormatRawCellContents(-12.343, -1, "##.##*)"));
            Assert.AreEqual(".", dfUS.FormatRawCellContents(0.0, -1, "##.##*)"));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.343, -1, "*-##.##"));
            Assert.AreEqual("-12.34", dfUS.FormatRawCellContents(-12.343, -1, "*-##.##"));
            Assert.AreEqual(".", dfUS.FormatRawCellContents(0.0, -1, "*-##.##"));
        }

        /**
         * Test that the special Excel month format MMMMM
         *  Gets turned into the first letter of the month
         */
        [Test]
        public void TestMMMMM()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));

            DateTime c = new DateTime(2010, 6, 1, 2, 0, 0, 0);
            System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            Assert.AreEqual("2010-J-1 2:00:00", dfUS.FormatRawCellContents(
                  DateUtil.GetExcelDate(c, false), -1, "YYYY-MMMMM-D h:mm:ss"
            ));
        }
        /**
         * Tests that we do AM/PM handling properly
         */
        [Test]
        public void TestAMPM()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));

            Assert.AreEqual("06:00", dfUS.FormatRawCellContents(0.25, -1, "hh:mm"));
            Assert.AreEqual("18:00", dfUS.FormatRawCellContents(0.75, -1, "hh:mm"));

            Assert.AreEqual("06:00 AM", dfUS.FormatRawCellContents(0.25, -1, "hh:mm AM/PM"));
            Assert.AreEqual("06:00 PM", dfUS.FormatRawCellContents(0.75, -1, "hh:mm AM/PM"));

            Assert.AreEqual("1904-01-01 06:00:00 AM", dfUS.FormatRawCellContents(0.25, -1, "yyyy-mm-dd hh:mm:ss AM/PM", true));
            Assert.AreEqual("1904-01-01 06:00:00 PM", dfUS.FormatRawCellContents(0.75, -1, "yyyy-mm-dd hh:mm:ss AM/PM", true));
        }
        /**
         * Test that we can handle elapsed time,
         *  eg formatting 1 day 4 hours as 28 hours
         */
        [Test]
        public void TestElapsedTime()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
            double hour = 1.0 / 24.0;

            Assert.AreEqual("01:00", dfUS.FormatRawCellContents(1 * hour, -1, "hh:mm"));
            Assert.AreEqual("05:00", dfUS.FormatRawCellContents(5 * hour, -1, "hh:mm"));
            Assert.AreEqual("20:00", dfUS.FormatRawCellContents(20 * hour, -1, "hh:mm"));
            Assert.AreEqual("23:00", dfUS.FormatRawCellContents(23 * hour, -1, "hh:mm"));
            Assert.AreEqual("00:00", dfUS.FormatRawCellContents(24 * hour, -1, "hh:mm"));
            Assert.AreEqual("02:00", dfUS.FormatRawCellContents(26 * hour, -1, "hh:mm"));
            Assert.AreEqual("20:00", dfUS.FormatRawCellContents(44 * hour, -1, "hh:mm"));
            Assert.AreEqual("02:00", dfUS.FormatRawCellContents(50 * hour, -1, "hh:mm"));

            Assert.AreEqual("01:00", dfUS.FormatRawCellContents(1 * hour, -1, "[hh]:mm"));
            Assert.AreEqual("05:00", dfUS.FormatRawCellContents(5 * hour, -1, "[hh]:mm"));
            Assert.AreEqual("20:00", dfUS.FormatRawCellContents(20 * hour, -1, "[hh]:mm"));
            Assert.AreEqual("23:00", dfUS.FormatRawCellContents(23 * hour, -1, "[hh]:mm"));
            Assert.AreEqual("24:00", dfUS.FormatRawCellContents(24 * hour, -1, "[hh]:mm"));
            Assert.AreEqual("26:00", dfUS.FormatRawCellContents(26 * hour, -1, "[hh]:mm"));
            Assert.AreEqual("44:00", dfUS.FormatRawCellContents(44 * hour, -1, "[hh]:mm"));
            Assert.AreEqual("50:00", dfUS.FormatRawCellContents(50 * hour, -1, "[hh]:mm"));

            Assert.AreEqual("01", dfUS.FormatRawCellContents(1 * hour, -1, "[hh]"));
            Assert.AreEqual("05", dfUS.FormatRawCellContents(5 * hour, -1, "[hh]"));
            Assert.AreEqual("20", dfUS.FormatRawCellContents(20 * hour, -1, "[hh]"));
            Assert.AreEqual("23", dfUS.FormatRawCellContents(23 * hour, -1, "[hh]"));
            Assert.AreEqual("24", dfUS.FormatRawCellContents(24 * hour, -1, "[hh]"));
            Assert.AreEqual("26", dfUS.FormatRawCellContents(26 * hour, -1, "[hh]"));
            Assert.AreEqual("44", dfUS.FormatRawCellContents(44 * hour, -1, "[hh]"));
            Assert.AreEqual("50", dfUS.FormatRawCellContents(50 * hour, -1, "[hh]"));

            double minute = 1.0 / 24.0 / 60.0;
            Assert.AreEqual("01:00", dfUS.FormatRawCellContents(1 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("05:00", dfUS.FormatRawCellContents(5 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("20:00", dfUS.FormatRawCellContents(20 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("23:00", dfUS.FormatRawCellContents(23 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("24:00", dfUS.FormatRawCellContents(24 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("26:00", dfUS.FormatRawCellContents(26 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("44:00", dfUS.FormatRawCellContents(44 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("50:00", dfUS.FormatRawCellContents(50 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("59:00", dfUS.FormatRawCellContents(59 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("60:00", dfUS.FormatRawCellContents(60 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("61:00", dfUS.FormatRawCellContents(61 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("119:00", dfUS.FormatRawCellContents(119 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("120:00", dfUS.FormatRawCellContents(120 * minute, -1, "[mm]:ss"));
            Assert.AreEqual("121:00", dfUS.FormatRawCellContents(121 * minute, -1, "[mm]:ss"));

            Assert.AreEqual("01", dfUS.FormatRawCellContents(1 * minute, -1, "[mm]"));
            Assert.AreEqual("05", dfUS.FormatRawCellContents(5 * minute, -1, "[mm]"));
            Assert.AreEqual("20", dfUS.FormatRawCellContents(20 * minute, -1, "[mm]"));
            Assert.AreEqual("23", dfUS.FormatRawCellContents(23 * minute, -1, "[mm]"));
            Assert.AreEqual("24", dfUS.FormatRawCellContents(24 * minute, -1, "[mm]"));
            Assert.AreEqual("26", dfUS.FormatRawCellContents(26 * minute, -1, "[mm]"));
            Assert.AreEqual("44", dfUS.FormatRawCellContents(44 * minute, -1, "[mm]"));
            Assert.AreEqual("50", dfUS.FormatRawCellContents(50 * minute, -1, "[mm]"));
            Assert.AreEqual("59", dfUS.FormatRawCellContents(59 * minute, -1, "[mm]"));
            Assert.AreEqual("60", dfUS.FormatRawCellContents(60 * minute, -1, "[mm]"));
            Assert.AreEqual("61", dfUS.FormatRawCellContents(61 * minute, -1, "[mm]"));
            Assert.AreEqual("119", dfUS.FormatRawCellContents(119 * minute, -1, "[mm]"));
            Assert.AreEqual("120", dfUS.FormatRawCellContents(120 * minute, -1, "[mm]"));
            Assert.AreEqual("121", dfUS.FormatRawCellContents(121 * minute, -1, "[mm]"));

            double second = 1.0 / 24.0 / 60.0 / 60.0;
            Assert.AreEqual("86400", dfUS.FormatRawCellContents(86400 * second, -1, "[ss]"));
            Assert.AreEqual("01", dfUS.FormatRawCellContents(1 * second, -1, "[ss]"));
            Assert.AreEqual("05", dfUS.FormatRawCellContents(5 * second, -1, "[ss]"));
            Assert.AreEqual("20", dfUS.FormatRawCellContents(20 * second, -1, "[ss]"));
            Assert.AreEqual("23", dfUS.FormatRawCellContents(23 * second, -1, "[ss]"));
            Assert.AreEqual("24", dfUS.FormatRawCellContents(24 * second, -1, "[ss]"));
            Assert.AreEqual("26", dfUS.FormatRawCellContents(26 * second, -1, "[ss]"));
            Assert.AreEqual("44", dfUS.FormatRawCellContents(44 * second, -1, "[ss]"));
            Assert.AreEqual("50", dfUS.FormatRawCellContents(50 * second, -1, "[ss]"));
            Assert.AreEqual("59", dfUS.FormatRawCellContents(59 * second, -1, "[ss]"));
            Assert.AreEqual("60", dfUS.FormatRawCellContents(60 * second, -1, "[ss]"));
            Assert.AreEqual("61", dfUS.FormatRawCellContents(61 * second, -1, "[ss]"));
            Assert.AreEqual("119", dfUS.FormatRawCellContents(119 * second, -1, "[ss]"));
            Assert.AreEqual("120", dfUS.FormatRawCellContents(120 * second, -1, "[ss]"));
            Assert.AreEqual("121", dfUS.FormatRawCellContents(121 * second, -1, "[ss]"));

            //boolean jdk_1_5 = System.getProperty("java.vm.version").startsWith("1.5");
            //if(!jdk_1_5) {
           // YK: the tests below were written under JDK 1.6 and assume that
           // the rounding mode in the underlying decimal formatters is HALF_UP
           // It is not so JDK 1.5 where the default rounding mode is HALV_EVEN and cannot be changed.


            Assert.AreEqual("27:18:08", dfUS.FormatRawCellContents(1.1376, -1, "[h]:mm:ss"));
            Assert.AreEqual("28:48:00", dfUS.FormatRawCellContents(1.2, -1, "[h]:mm:ss"));
            Assert.AreEqual("29:31:12", dfUS.FormatRawCellContents(1.23, -1, "[h]:mm:ss"));
            Assert.AreEqual("31:26:24", dfUS.FormatRawCellContents(1.31, -1, "[h]:mm:ss"));

            Assert.AreEqual("27:18:08", dfUS.FormatRawCellContents(1.1376, -1, "[hh]:mm:ss"));
            Assert.AreEqual("28:48:00", dfUS.FormatRawCellContents(1.2, -1, "[hh]:mm:ss"));
            Assert.AreEqual("29:31:12", dfUS.FormatRawCellContents(1.23, -1, "[hh]:mm:ss"));
            Assert.AreEqual("31:26:24", dfUS.FormatRawCellContents(1.31, -1, "[hh]:mm:ss"));

            Assert.AreEqual("57:07.2", dfUS.FormatRawCellContents(.123, -1, "mm:ss.0;@"));
            Assert.AreEqual("57:41.8", dfUS.FormatRawCellContents(.1234, -1, "mm:ss.0;@"));
            Assert.AreEqual("57:41.76", dfUS.FormatRawCellContents(.1234, -1, "mm:ss.00;@"));
            Assert.AreEqual("57:41.760", dfUS.FormatRawCellContents(.1234, -1, "mm:ss.000;@"));
            Assert.AreEqual("24:00.0", dfUS.FormatRawCellContents(123456.6, -1, "mm:ss.0"));

            //}
        }
        [Test]
        public void TestDateWindowing()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
            Assert.AreEqual("1899-12-31 00:00:00", dfUS.FormatRawCellContents(0.0, -1, "yyyy-mm-dd hh:mm:ss"));
            Assert.AreEqual("1899-12-31 00:00:00", dfUS.FormatRawCellContents(0.0, -1, "yyyy-mm-dd hh:mm:ss", false));
            Assert.AreEqual("1904-01-01 00:00:00", dfUS.FormatRawCellContents(0.0, -1, "yyyy-mm-dd hh:mm:ss", true));
        }
        [Test]
        public void TestScientificNotation()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));

            Assert.AreEqual("1.23E+01", dfUS.FormatRawCellContents(12.343, -1, "0.00E+00"));
            Assert.AreEqual("-1.23E+01", dfUS.FormatRawCellContents(-12.343, -1, "0.00E+00"));
            Assert.AreEqual("0.00E+00", dfUS.FormatRawCellContents(0.0, -1, "0.00E+00"));
        }
        [Test]
        public void TestInvalidDate()
        {
            //DataFormatter df1 = new DataFormatter(Locale.US);
            DataFormatter df1 = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
            //Assert.AreEqual("-1.0", df1.FormatRawCellContents(-1, -1, "mm/dd/yyyy"));
            //in java -1.toString() is "-1.0", but in C# -1.ToString() is "-1".
            Assert.AreEqual("-1", df1.FormatRawCellContents(-1, -1, "mm/dd/yyyy"));
            //DataFormatter df2 = new DataFormatter(Locale.US);
            DataFormatter df2 = new DataFormatter(CultureInfo.GetCultureInfo("en-US"), true);
            Assert.AreEqual("###############################################################################################################################################################################################################################################################",
                    df2.FormatRawCellContents(-1, -1, "mm/dd/yyyy"));
        }
        [Test]
        public void TestEscapes()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
            Assert.AreEqual("1901-01-01", dfUS.FormatRawCellContents(367.0, -1, "yyyy-mm-dd"));
            Assert.AreEqual("1901-01-01", dfUS.FormatRawCellContents(367.0, -1, "yyyy\\-mm\\-dd"));

            Assert.AreEqual("1901.01.01", dfUS.FormatRawCellContents(367.0, -1, "yyyy.mm.dd"));
            Assert.AreEqual("1901.01.01", dfUS.FormatRawCellContents(367.0, -1, "yyyy\\.mm\\.dd"));

            Assert.AreEqual("1901/01/01", dfUS.FormatRawCellContents(367.0, -1, "yyyy/mm/dd"));
            Assert.AreEqual("1901/01/01", dfUS.FormatRawCellContents(367.0, -1, "yyyy\\/mm\\/dd"));
        }
        [Test]
        public void TestFormatsWithPadding()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"), true);

            // These request space-padding, based on the cell width
            // There should always be one space after, variable (non-zero) amount before
            // Because the Cell Width isn't available, this gets emulated with
            //  4 leading spaces, or a minus then 3 leading spaces
            // This isn't all that consistent, but it's the best we can really manage...
            Assert.AreEqual("    1,234.56 ", dfUS.FormatRawCellContents(1234.56, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            Assert.AreEqual("-   1,234.56 ", dfUS.FormatRawCellContents(-1234.56, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            Assert.AreEqual("    12.34 ", dfUS.FormatRawCellContents(12.34, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            Assert.AreEqual("-   12.34 ", dfUS.FormatRawCellContents(-12.34, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));

            Assert.AreEqual("    0.10 ", dfUS.FormatRawCellContents(0.1, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            Assert.AreEqual("-   0.10 ", dfUS.FormatRawCellContents(-0.1, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            // TODO Fix this, we are randomly adding a 0 at the end that souldn't be there
            //Assert.AreEqual("     -   ", dfUS.FormatRawCellContents(0.0, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));

            Assert.AreEqual(" $   1.10 ", dfUS.FormatRawCellContents(1.1, -1, "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-"));
            Assert.AreEqual("-$   1.10 ", dfUS.FormatRawCellContents(-1.1, -1, "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-"));
            // TODO Fix this, we are randomly adding a 0 at the end that souldn't be there
            //Assert.AreEqual(" $    -   ", dfUS.FormatRawCellContents( 0.0, -1, "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-"));
        }
        [Test]
        public void TestErrors()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"), true);

            // Create a spreadsheet with some formula errors in it
            IWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(0, CellType.Error);

            c.SetCellErrorValue(FormulaError.DIV0.Code);
            Assert.AreEqual(FormulaError.DIV0.String, dfUS.FormatCellValue(c));

            c.SetCellErrorValue(FormulaError.REF.Code);
            Assert.AreEqual(FormulaError.REF.String, dfUS.FormatCellValue(c));
        }

        [Test]
        public void TestBoolean()
        {
            DataFormatter formatter = new DataFormatter();
            // Create a spreadsheet with some TRUE/FALSE boolean values in1 it
            IWorkbook wb = new HSSFWorkbook();
            try
            {
                ISheet s = wb.CreateSheet();
                IRow r = s.CreateRow(0);
                ICell c = r.CreateCell(0);
                c.SetCellValue(true);
                Assert.AreEqual("TRUE", formatter.FormatCellValue(c));
                c.SetCellValue(false);
                Assert.AreEqual("FALSE", formatter.FormatCellValue(c));
            }
            finally
            {
                wb.Close();
            }
        }

        /**
         * While we don't currently support using a locale code at
         *  the start of a format string to format it differently, we
         *  should at least handle it as it if wasn't there
         */
        [Test]
        public void TestDatesWithLocales()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"), true);

            String dateFormatEnglish = "[$-409]mmmm dd yyyy  h:mm AM/PM";
            String dateFormatChinese = "[$-804]mmmm dd yyyy  h:mm AM/PM";

            // Check we format the English one correctly
            double date = 26995.477777777778;
            Assert.AreEqual(
                    "November 27 1973  11:28 AM",
                    dfUS.FormatRawCellContents(date, -1, dateFormatEnglish)
            );

            // Check that, in the absence of locale support, we handle
            //  the Chinese one the same as the English one
            Assert.AreEqual(
                    "November 27 1973  11:28 AM",
                    dfUS.FormatRawCellContents(date, -1, dateFormatChinese)
            );
        }

        //TODO Fix these so that they work
        [Test]
        [Ignore("Fix these so that they work")]
        public void TestCustomFormats()
        {
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"), true);
            String fmt;

            fmt = "\"At\" H:MM AM/PM \"on\" DDDD MMMM D\",\" YYYY";
            Assert.AreEqual(
                  "At 4:20 AM on Thursday May 17, 2007",
                  dfUS.FormatRawCellContents(39219.1805636921, -1, fmt)
            );

            fmt = "0 \"dollars and\" .00 \"cents\"";
            Assert.AreEqual("19 dollars and .99 cents", dfUS.FormatRawCellContents(19.99, -1, fmt));
        }
        /**
         * ExcelStyleDateFormatter should work for Milliseconds too
         */
        [Test]
        public void TestExcelStyleDateFormatterStringOnMillis()
        {
            // Test directly with the .000 style
            SimpleDateFormat formatter1 = new ExcelStyleDateFormatter("ss.000");
            CultureInfo culture = CultureInfo.GetCultureInfo("en-US");
            DateTime dt = DateTime.Now.Date;
            Assert.AreEqual("00.001", formatter1.Format(dt.AddMilliseconds(1L), culture));
            Assert.AreEqual("00.010", formatter1.Format(dt.AddMilliseconds(10L), culture));
            Assert.AreEqual("00.100", formatter1.Format(dt.AddMilliseconds(100L), culture));
            Assert.AreEqual("01.000", formatter1.Format(dt.AddMilliseconds(1000L), culture));
            Assert.AreEqual("01.001", formatter1.Format(dt.AddMilliseconds(1001L), culture));
            Assert.AreEqual("10.000", formatter1.Format(dt.AddMilliseconds(10000L), culture));
            Assert.AreEqual("10.001", formatter1.Format(dt.AddMilliseconds(10001L), culture));

            // Test directly with the .SSS style
            SimpleDateFormat formatter2 = new ExcelStyleDateFormatter("ss.fff");

            Assert.AreEqual("00.001", formatter2.Format(dt.AddMilliseconds(1L), culture));
            Assert.AreEqual("00.010", formatter2.Format(dt.AddMilliseconds(10L), culture));
            Assert.AreEqual("00.100", formatter2.Format(dt.AddMilliseconds(100L), culture));
            Assert.AreEqual("01.000", formatter2.Format(dt.AddMilliseconds(1000L), culture));
            Assert.AreEqual("01.001", formatter2.Format(dt.AddMilliseconds(1001L), culture));
            Assert.AreEqual("10.000", formatter2.Format(dt.AddMilliseconds(10000L), culture));
            Assert.AreEqual("10.001", formatter2.Format(dt.AddMilliseconds(10001L), culture));


            // Test via DataFormatter
            DataFormatter dfUS = new DataFormatter(culture, true);
            Assert.AreEqual("01.010", dfUS.FormatRawCellContents(0.0000116898, -1, "ss.000"));
        }
        [Test]
        public void TestBug54786()
        {
            DataFormatter formatter = new DataFormatter();
            String format = "[h]\"\"h\"\" m\"\"m\"\"";
            Assert.IsTrue(DateUtil.IsADateFormat(-1, format));
            Assert.IsTrue(DateUtil.IsValidExcelDate(_15_MINUTES));

            Assert.AreEqual("1h 0m", formatter.FormatRawCellContents(_15_MINUTES, -1, format, false));
            Assert.AreEqual("0.041666667", formatter.FormatRawCellContents(_15_MINUTES, -1, "[h]'h'", false));
            Assert.AreEqual("1h 0m\"", formatter.FormatRawCellContents(_15_MINUTES, -1, "[h]\"\"h\"\" m\"\"m\"\"\"", false));
            Assert.AreEqual("1h", formatter.FormatRawCellContents(_15_MINUTES, -1, "[h]\"\"h\"\"", false));
            Assert.AreEqual("h1", formatter.FormatRawCellContents(_15_MINUTES, -1, "\"\"h\"\"[h]", false));
            Assert.AreEqual("h1", formatter.FormatRawCellContents(_15_MINUTES, -1, "\"\"h\"\"h", false));
            Assert.AreEqual(" 60", formatter.FormatRawCellContents(_15_MINUTES, -1, " [m]", false));
            Assert.AreEqual("h60", formatter.FormatRawCellContents(_15_MINUTES, -1, "\"\"h\"\"[m]", false));
            Assert.AreEqual("m1", formatter.FormatRawCellContents(_15_MINUTES, -1, "\"\"m\"\"h", false));

            try
            {
                Assert.AreEqual("1h 0m\"", formatter.FormatRawCellContents(_15_MINUTES, -1, "[h]\"\"h\"\" m\"\"m\"\"\"\"", false));
                Assert.Fail("Catches exception because of invalid format, i.e. trailing quoting");
            }
            catch (Exception)
            {
                //Assert.IsTrue(e.Message.Contains("Cannot format given Object as a Number"));
            }
        }
        [Test]
        public void TestIsADateFormat()
        {
            // first check some cases that should not be a date, also call multiple times to ensure the cache is used
            Assert.IsFalse(DateUtil.IsADateFormat(-1, null));
            Assert.IsFalse(DateUtil.IsADateFormat(-1, null));
            Assert.IsFalse(DateUtil.IsADateFormat(123, null));
            Assert.IsFalse(DateUtil.IsADateFormat(123, ""));
            Assert.IsFalse(DateUtil.IsADateFormat(124, ""));
            Assert.IsFalse(DateUtil.IsADateFormat(-1, ""));
            Assert.IsFalse(DateUtil.IsADateFormat(-1, ""));
            Assert.IsFalse(DateUtil.IsADateFormat(-1, "nodateformat"));

            // then also do the same for some valid date formats
            Assert.IsTrue(DateUtil.IsADateFormat(0x0e, null));
            Assert.IsTrue(DateUtil.IsADateFormat(0x2f, null));
            Assert.IsTrue(DateUtil.IsADateFormat(-1, "yyyy"));
            Assert.IsTrue(DateUtil.IsADateFormat(-1, "yyyy"));
            Assert.IsTrue(DateUtil.IsADateFormat(-1, "dd/mm/yy;[red]dd/mm/yy"));
            Assert.IsTrue(DateUtil.IsADateFormat(-1, "dd/mm/yy;[red]dd/mm/yy"));
            Assert.IsTrue(DateUtil.IsADateFormat(-1, "[h]"));
        }

        [Test]
        public void testLargeNumbersAndENotation()
        {
            assertFormatsTo("1E+86", 99999999999999999999999999999999999999999999999999999999999999999999999999999999999999d);
            assertFormatsTo("1E-84", 0.000000000000000000000000000000000000000000000000000000000000000000000000000000000001d);
            // Smallest double
            //assertFormatsTo("1E-323", 0.00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000001d);
            // "up to 11 numeric characters, with the decimal point counting as a numeric character"
            // https://support.microsoft.com/en-us/kb/65903
            assertFormatsTo("12345678911", 12345678911d);
            assertFormatsTo("1.23457E+11", 123456789112d);  // 12th digit of integer -> scientific
            assertFormatsTo("-12345678911", -12345678911d);
            assertFormatsTo("-1.23457E+11", -123456789112d);
            assertFormatsTo("0.1", 0.1);
            assertFormatsTo("0.000000001", 0.000000001);
            assertFormatsTo("1E-10", 0.0000000001);  // 12th digit
            assertFormatsTo("-0.000000001", -0.000000001);
            assertFormatsTo("-1E-10", -0.0000000001);
            assertFormatsTo("123.4567892", 123.45678919);  // excess decimals are simply rounded away
            assertFormatsTo("-123.4567892", -123.45678919);
            assertFormatsTo("1.234567893", 1.2345678925);  // rounding mode is half-up
            assertFormatsTo("-1.234567893", -1.2345678925);
            assertFormatsTo("1.23457E+19", 12345650000000000000d);
            assertFormatsTo("-1.23457E+19", -12345650000000000000d);
            assertFormatsTo("1.23457E-19", 0.0000000000000000001234565d);
            assertFormatsTo("-1.23457E-19", -0.0000000000000000001234565d);
            assertFormatsTo("1.000000001", 1.000000001);
            assertFormatsTo("1", 1.0000000001);
            assertFormatsTo("1234.567891", 1234.567891123456789d);
            assertFormatsTo("1234567.891", 1234567.891123456789d);
            assertFormatsTo("12345678912", 12345678911.63456789d);  // integer portion uses all 11 digits
            assertFormatsTo("12345678913", 12345678912.5d);  // half-up here too
            assertFormatsTo("-12345678913", -12345678912.5d);
            assertFormatsTo("1.23457E+11", 123456789112.3456789d);
        }
        private static void assertFormatsTo(String expected, double input)
        {
            IWorkbook wb = new HSSFWorkbook();
            try
            {
                ISheet s1 = wb.CreateSheet();
                IRow row = s1.CreateRow(0);
                ICell rawValue = row.CreateCell(0);
                rawValue.SetCellValue(input);
                ICellStyle newStyle = wb.CreateCellStyle();
                IDataFormat dataFormat = wb.CreateDataFormat();
                newStyle.DataFormat = (dataFormat.GetFormat("General"));
                String actual = new DataFormatter().FormatCellValue(rawValue);
                Assert.AreEqual(expected, actual);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestFormulaEvaluation()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("FormulaEvalTestData.xls");
            CellReference ref1 = new CellReference("D47");
            ICell cell = wb.GetSheetAt(0).GetRow(ref1.Row).GetCell(ref1.Col);
            //noinspection deprecation
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual("G9:K9 I7:I12", cell.CellFormula);
            DataFormatter formatter = new DataFormatter();
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual("5.6789", formatter.FormatCellValue(cell, evaluator));
            wb.Close();
        }

        /**
         * bug 60031: DataFormatter parses months incorrectly when put at the end of date segment
         */
        [Test]
        public void TestBug60031()
        {
            // 23-08-2016 08:51:01 which is 42605.368761574071 as double will be parsed
            // with format "yyyy-dd-MM HH:mm:ss" into "2016-23-51 08:51:01".
            DataFormatter dfUS = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
            Assert.AreEqual("2016-23-08 08:51:01", dfUS.FormatRawCellContents(42605.368761574071, -1, "yyyy-dd-MM HH:mm:ss"));
            Assert.AreEqual("2016-23 08:51:01 08", dfUS.FormatRawCellContents(42605.368761574071, -1, "yyyy-dd HH:mm:ss MM"));
            Assert.AreEqual("2017-12-01 January 09:54:33", dfUS.FormatRawCellContents(42747.412892397523, -1, "yyyy-dd-MM MMMM HH:mm:ss"));

            Assert.AreEqual("08", dfUS.FormatRawCellContents(42605.368761574071, -1, "MM"));
            Assert.AreEqual("01", dfUS.FormatRawCellContents(42605.368761574071, -1, "ss"));

            // From Excel help:
            /*
                The "m" or "mm" code must appear immediately after the "h" or"hh"
                code or immediately before the "ss" code; otherwise, Microsoft
                Excel displays the month instead of minutes."
              */
            Assert.AreEqual("08", dfUS.FormatRawCellContents(42605.368761574071, -1, "mm"));
            Assert.AreEqual("08:51", dfUS.FormatRawCellContents(42605.368761574071, -1, "hh:mm"));
            Assert.AreEqual("51:01", dfUS.FormatRawCellContents(42605.368761574071, -1, "mm:ss"));
        }
    }
}