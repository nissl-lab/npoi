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
namespace TestCases.SS.Format
{
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;

    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Format;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using TestCases.SS;
    using System.Diagnostics;
    using SixLabors.ImageSharp;

    /**
     * This class is a base class for spreadsheet-based tests, such as are used for
     * cell formatting.  This Reads tests from the spreadsheet, as well as Reading
     * flags that can be used to paramterize these tests.
     * <p/>
     * Each test has four parts: The expected result (column A), the format string
     * (column B), the value to format (column C), and a comma-Separated list of
     * categores that this test falls in1. Normally all tests are Run, but if the
     * flag "Categories" is not empty, only tests that have at least one category
     * listed in "Categories" are Run.
     */
    //[TestFixture]
    public class CellFormatTestBase
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(CellFormatTestBase));

        private ITestDataProvider _testDataProvider;

        protected IWorkbook workbook;

        private String testFile;
        private Dictionary<String, String> testFlags;
        private bool tryAllColors;
        //private Label label;

        private static String[] COLOR_NAMES =
            {"Black", "Red", "Green", "Blue", "Yellow", "Cyan", "Magenta",
                    "White"};
        private static Color[] COLORS = { Color.Black, Color.Red, Color.Green, Color.Blue, Color.Yellow, Color.Cyan, Color.Magenta, Color.Wheat };

        public static Color TEST_COLOR = Color.Orange; //Darker();

        protected CellFormatTestBase(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }

        public abstract class CellValue
        {
            public abstract Object GetValue(ICell cell);

            public Color GetColor(ICell cell)
            {
                return TEST_COLOR;
            }

            public virtual void Equivalent(String expected, String actual, CellFormatPart format)
            {
                Assert.AreEqual('"' + expected + '"',
                        '"' + actual + '"', "format \"" + format.ToString() + "\"");
            }
        }

        protected void RunFormatTests(String workbookName, CellValue valueGetter)
        {

            OpenWorkbook(workbookName);

            ReadFlags(workbook);

            SortedList<string, object> runCategories = new SortedList<string, object>(StringComparer.InvariantCultureIgnoreCase);
            String RunCategoryList = flagString("Categories", "");
            Regex regex = new Regex("\\s*,\\s*");
            if (RunCategoryList != null)
            {
                foreach (string s in regex.Split(RunCategoryList))
                    if (!runCategories.ContainsKey(s))
                        runCategories.Add(s, null);
                runCategories.Remove(""); // this can be found and means nothing
            }

            ISheet sheet = workbook.GetSheet("Tests");
            int end = sheet.LastRowNum;
            // Skip the header row, therefore "+ 1"
            for (int r = sheet.FirstRowNum + 1; r <= end; r++)
            {
                IRow row = sheet.GetRow(r);
                if (row == null)
                    continue;
                int cellnum = 0;
                String expectedText = row.GetCell(cellnum).StringCellValue;
                String format = row.GetCell(1).StringCellValue;
                String testCategoryList = row.GetCell(3).StringCellValue;
                bool byCategory = RunByCategory(runCategories, testCategoryList);
                if ((expectedText.Length > 0 || format.Length > 0) && byCategory)
                {
                    ICell cell = row.GetCell(2);
                    Debug.WriteLine(string.Format("expectedText: {0}, format:{1}", expectedText, format));
                    if (format == "hh:mm:ss a/p")
                        expectedText = expectedText.ToUpper();
                    else if (format == "H:M:S.00 a/p")
                        expectedText = expectedText.ToUpper();
                    tryFormat(r, expectedText, format, valueGetter, cell);
                }
            }
        }

        /**
         * Open a given workbook.
         *
         * @param workbookName The workbook name.  This is presumed to live in the
         *                     "spreadsheets" directory under the directory named in
         *                     the Java property "POI.testdata.path".
         *
         * @throws IOException
         */
        protected void OpenWorkbook(String workbookName)
        {
            workbook = _testDataProvider.OpenSampleWorkbook(workbookName);
            workbook.MissingCellPolicy = MissingCellPolicy.CREATE_NULL_AS_BLANK;//Row.CREATE_NULL_AS_BLANK);
            testFile = workbookName;
        }

        /**
         * Read the flags from the workbook.  Flags are on the sheet named "Flags",
         * and consist of names in column A and values in column B.  These are Put
         * into a map that can be queried later.
         *
         * @param wb The workbook to look in1.
         */
        private void ReadFlags(IWorkbook wb)
        {
            ISheet flagSheet = wb.GetSheet("Flags");
            testFlags = new Dictionary<String, String>();
            if (flagSheet != null)
            {
                int end = flagSheet.LastRowNum;
                // Skip the header row, therefore "+ 1"
                for (int r = flagSheet.FirstRowNum + 1; r <= end; r++)
                {
                    IRow row = flagSheet.GetRow(r);
                    if (row == null)
                        continue;
                    String flagName = row.GetCell(0).StringCellValue;
                    String flagValue = row.GetCell(1).StringCellValue;
                    if (flagName.Length > 0)
                    {
                        testFlags.Add(flagName, flagValue);
                    }
                }
            }

            tryAllColors = flagBoolean("AllColors", true);
        }

        /**
         * Returns <tt>true</tt> if any of the categories for this run are Contained
         * in the test's listed categories.
         *
         * @param categories     The categories of tests to be Run.  If this is
         *                       empty, then all tests will be Run.
         * @param testCategories The categories that this test is in1.  This is a
         *                       comma-Separated list.  If <em>any</em> tests in
         *                       this list are in <tt>categories</tt>, the test will
         *                       be Run.
         *
         * @return <tt>true</tt> if the test should be Run.
         */
        private bool RunByCategory(SortedList<string, object> categories,
                String testCategories)
        {

            if (categories.Count == 0)
                return true;
            // If there are specified categories, find out if this has one of them
            Regex regex = new Regex("\\s*,\\s*");

            foreach (String category in regex.Split(testCategories))//.Split("\\s*,\\s*"))
            {
                if (categories.ContainsKey(category))
                {
                    return true;
                }
            }
            return false;
        }

        Color labelForeColor;
        private void tryFormat(int row, String expectedText, String desc,
                CellValue getter, ICell cell)
        {

            Object value = getter.GetValue(cell);
            Color testColor = getter.GetColor(cell);

            labelForeColor = testColor;

            logger.Log(POILogger.INFO, String.Format("Row %d: \"%s\" -> \"%s\": expected \"%s\"", row + 1,
                    value.ToString(), desc, expectedText));

            String actualText = tryColor(desc, null, getter, value, expectedText,
                    testColor);
            logger.Log(POILogger.INFO, String.Format(", actual \"%s\")%n", actualText));

            if (tryAllColors && testColor != TEST_COLOR)
            {
                for (int i = 0; i < COLOR_NAMES.Length; i++)
                {
                    String cname = COLOR_NAMES[i];
                    tryColor(desc, cname, getter, value, expectedText, COLORS[i]);
                }
            }
        }

        private String tryColor(String desc, String cname, CellValue getter,
                Object value, String expectedText, Color expectedColor)
        {

            if (cname != null)
                desc = "[" + cname + "]" + desc;
            Color origColor = labelForeColor;
            CellFormatPart format = new CellFormatPart(desc);
            CellFormatResult result = format.Apply(value);
            if (!result.Applies)
            {
                // If this doesn't Apply, no color change is expected
                expectedColor = origColor;
            }

            String actualText = result.Text;
            Color actualColor = labelForeColor;
            getter.Equivalent(expectedText, actualText, format);
            Assert.AreEqual(
                    expectedColor, actualColor, cname == null ? "no color" : "color " + cname);
            return actualText;
        }
        /// <summary>
        ///  Returns the value for the given flag.  The flag has the value of <tt>true</tt> if the text value is <tt>"true"</tt>, <tt>"yes"</tt>, or <tt>"on"</tt> (ignoring case).
        /// </summary>
        /// <param name="flagName">The name of the flag to fetch.</param>
        /// <param name="expected">
        /// The value for the flag that is expected when the tests are run for a full test.  If the current value is not the expected one, 
        /// you will get a warning in the test output. This is so that you do not accidentally leave a flag set to a value that prevents Running some tests, thereby
        /// letting you accidentally release code that is not fully tested.
        /// </param>
        /// <returns></returns>
        protected bool flagBoolean(String flagName, bool expected)
        {
            String value = testFlags[(flagName)];
            bool isSet;
            if (value == null)
                isSet = false;
            else
            {
                isSet = value.Equals("true", StringComparison.InvariantCultureIgnoreCase) || value.Equals(
                        "yes", StringComparison.InvariantCultureIgnoreCase) || value.Equals("on", StringComparison.InvariantCultureIgnoreCase);
            }
            warnIfUnexpected(flagName, expected, isSet);
            return isSet;
        }

        /**
         * Returns the value for the given flag.
         *
         * @param flagName The name of the flag to fetch.
         * @param expected The value for the flag that is expected when the tests
         *                 are run for a full test.  If the current value is not the
         *                 expected one, you will Get a warning in the test output.
         *                 This is so that you do not accidentally leave a flag Set
         *                 to a value that prevents Running some tests, thereby
         *                 letting you accidentally release code that is not fully
         *                 tested.
         *
         * @return The value for the flag.
         */
        protected String flagString(String flagName, String expected)
        {
            String value = testFlags[(flagName)];
            if (value == null)
                value = "";
            warnIfUnexpected(flagName, expected, value);
            return value;
        }

        private void warnIfUnexpected(String flagName, Object expected,
                Object actual)
        {
            if (!actual.Equals(expected))
            {
                System.Console.WriteLine(
                        "WARNING: " + testFile + ": " + "Flag " + flagName +
                                " = \"" + actual + "\" [not \"" + expected + "\"]");
            }
        }
    }

}