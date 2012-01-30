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
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /**
     * Tests of {@link DataFormatter}
     *
     * See {@link TestHSSFDataFormatter} too for
     *  more Tests.
     */
    [TestClass]
    public class TestDataFormatter
    {
        /**
         * Test that we use the specified locale when deciding
         *   how to format normal numbers
         */
        [TestMethod]
        public void TestLocale()
        {
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));

            Assert.AreEqual("1234", dfUS.FormatRawCellContents(1234, -1, "@"));
            Assert.AreEqual("12.34", dfUS.FormatRawCellContents(12.34, -1, "@"));

            DataFormatter dfFR = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("fr-FR"));

            Assert.AreEqual("1234", dfFR.FormatRawCellContents(1234, -1, "@"));
            Assert.AreEqual("12,34", dfFR.FormatRawCellContents(12.34, -1, "@"));
        }

        /**
         * Ensure that colours Get correctly
         *  zapped from within the format strings
         */
        [TestMethod]
        public void TestColours()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));

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
        [TestMethod]
        public void TestColoursAndBrackets()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));

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

        /**
         * Test how we handle negative and zeros.
         * Note - some Tests are disabled as DecimalFormat
         *  and Excel differ, and workarounds are not
         *  yet in place for all of these
         */
        [TestMethod]
        public void TestNegativeZero()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));

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
         * Test that _x (blank with the space taken by "x")
         *  and *x (fill to the column width with "x"s) are
         *  correctly ignored by us.
         */
        [TestMethod]
        public void TestPAddingSpaces()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));
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
        [TestMethod]
        public void TestPAddingSpacesCSV()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"), true);
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
        [TestMethod]
        public void TestMMMMM()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));

            DateTime c = new DateTime(2010, 6, 1, 2, 0, 0, 0);
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("en-US");
            Assert.AreEqual("2010-J-1 2:00:00", dfUS.FormatRawCellContents(
                  DateUtil.GetExcelDate(c, false), -1, "YYYY-MMMMM-D h:mm:ss"
            ));
        }

        /**
         * Test that we can handle elapsed time,
         *  eg formatting 1 day 4 hours as 28 hours
         */
        [TestMethod]
        public void TestElapsedTime()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));
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
        }
        [TestMethod]
        public void TestDateWindowing()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));
            Assert.AreEqual("1899-12-31 00:00:00", dfUS.FormatRawCellContents(0.0, -1, "yyyy-mm-dd hh:mm:ss"));
            Assert.AreEqual("1899-12-31 00:00:00", dfUS.FormatRawCellContents(0.0, -1, "yyyy-mm-dd hh:mm:ss", false));
            Assert.AreEqual("1904-01-01 00:00:00", dfUS.FormatRawCellContents(0.0, -1, "yyyy-mm-dd hh:mm:ss", true));
        }
        [TestMethod]
        public void TestScientificNotation()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));

            Assert.AreEqual("1.23E+01", dfUS.FormatRawCellContents(12.343, -1, "0.00E+00"));
            Assert.AreEqual("-1.23E+01", dfUS.FormatRawCellContents(-12.343, -1, "0.00E+00"));
            Assert.AreEqual("0.00E+00", dfUS.FormatRawCellContents(0.0, -1, "0.00E+00"));
        }
        [TestMethod]
        public void TestInvalidDate()
        {
            //DataFormatter df1 = new DataFormatter(Locale.US);
            DataFormatter df1 = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));
            //Assert.AreEqual("-1.0", df1.FormatRawCellContents(-1, -1, "mm/dd/yyyy"));
            //in java -1.toString() is "-1.0", but in C# -1.ToString() is "-1".
            Assert.AreEqual("-1", df1.FormatRawCellContents(-1, -1, "mm/dd/yyyy"));
            //DataFormatter df2 = new DataFormatter(Locale.US);
            DataFormatter df2 = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"), true);
            Assert.AreEqual("###############################################################################################################################################################################################################################################################",
                    df2.FormatRawCellContents(-1, -1, "mm/dd/yyyy"));
        }
        [TestMethod]
        public void TestEscapes()
        {
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"));
            Assert.AreEqual("1901-01-01", dfUS.FormatRawCellContents(367.0, -1, "yyyy-mm-dd"));
            Assert.AreEqual("1901-01-01", dfUS.FormatRawCellContents(367.0, -1, "yyyy\\-mm\\-dd"));

            Assert.AreEqual("1901.01.01", dfUS.FormatRawCellContents(367.0, -1, "yyyy.mm.dd"));
            Assert.AreEqual("1901.01.01", dfUS.FormatRawCellContents(367.0, -1, "yyyy\\.mm\\.dd"));

            Assert.AreEqual("1901/01/01", dfUS.FormatRawCellContents(367.0, -1, "yyyy/mm/dd"));
            Assert.AreEqual("1901/01/01", dfUS.FormatRawCellContents(367.0, -1, "yyyy\\/mm\\/dd"));
        }
        [TestMethod]
        public void TestOther()
        {
            //DataFormatter dfUS = new DataFormatter(Locale.US, true);
            DataFormatter dfUS = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("en-US"), true);
            Assert.AreEqual(" 12.34 ", dfUS.FormatRawCellContents(12.34, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            Assert.AreEqual("-12.34 ", dfUS.FormatRawCellContents(-12.34, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            Assert.AreEqual(" -   ", dfUS.FormatRawCellContents(0.0, -1, "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-"));
            Assert.AreEqual(" $-   ", dfUS.FormatRawCellContents(0.0, -1, "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-"));
        }
    }

}