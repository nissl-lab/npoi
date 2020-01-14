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
    using NPOI.SS.Format;
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NUnit.Framework;
    using System;
    using System.Globalization;
    using System.Text;
    using System.Text.RegularExpressions;




    /** Test the individual CellFormatPart types. */
    [TestFixture]
    public class TestCellFormatPart : CellFormatTestBase
    {
        private static Regex NUMBER_EXTRACT_FMT = new Regex(
                "([-+]?[0-9]+)(\\.[0-9]+)?.*(?:(e).*?([+-]?[0-9]+))",
                RegexOptions.IgnoreCase);

        public TestCellFormatPart()
            : base(XSSFITestDataProvider.instance)
        {
        }
        private class CellValue1 : CellValue
        {
            public override object GetValue(ICell cell)
            {
                CellType type = CellFormat.UltimateType(cell);
                if (type == CellType.Boolean)
                    return cell.BooleanCellValue ? "TRUE" : "FALSE";
                else if (type == CellType.Numeric)
                    return cell.NumericCellValue;
                else
                    return cell.StringCellValue;
            }
        }
        [Test]
        public void TestGeneralFormat()
        {
            RunFormatTests("GeneralFormatTests.xlsx", new CellValue1());
        }
        private class CellValue2 : CellValue
        {
            public override object GetValue(ICell cell)
            {
                return cell.NumericCellValue;
            }
        }
        [Test]
        public void TestNumberFormat()
        {
            RunFormatTests("NumberFormatTests.xlsx", new CellValue2());
        }
        private class CellValue3 : CellValue
        {
            public override object GetValue(ICell cell)
            {
                return cell.NumericCellValue;
            }
            public override void Equivalent(String expected, String actual,
                   CellFormatPart format)
            {
                double expectedVal = ExtractNumber(expected);
                double actualVal = ExtractNumber(actual);
                // equal within 1%
                double delta = expectedVal / 100;
                Assert.AreEqual(expectedVal, actualVal, delta, "format \"" + format + "\"," + expected + " ~= " +
                        actual);
            }
        }
        [Test]
        public void TestNumberApproxFormat()
        {
            RunFormatTests("NumberFormatApproxTests.xlsx", new CellValue3());
        }
        private class CellValue4 : CellValue
        {

            public override object GetValue(ICell cell)
            {
                return cell.DateCellValue;
            }
        }
        [Test]
        public void TestDateFormat()
        {
            RunFormatTests("DateFormatTests.xlsx", new CellValue4());
        }

        [Test]
        public void TestElapsedFormat()
        {
            RunFormatTests("ElapsedFormatTests.xlsx", new CellValue2());
        }
        private class CellValue6 : CellValue
        {
            public override object GetValue(ICell cell)
            {
                if (CellFormat.UltimateType(cell) == CellType.Boolean)
                    return cell.BooleanCellValue ? "TRUE" : "FALSE";
                else
                    return cell.StringCellValue;
            }
        }
        [Test]
        public void TestTextFormat()
        {
            RunFormatTests("TextFormatTests.xlsx", new CellValue6());
        }

        [Test]
        public void TestConditions()
        {
            RunFormatTests("FormatConditionTests.xlsx", new CellValue2());
        }

        private static double ExtractNumber(String str)
        {
            Match m = NUMBER_EXTRACT_FMT.Match(str);
            if (!m.Success)
                throw new ArgumentException(
                        "Cannot find numer in \"" + str + "\"");

            StringBuilder sb = new StringBuilder();
            // The groups in the pattern are the parts of the number
            for (int i = 1; i <= m.Groups.Count; i++)
            {
                String part = m.Groups[i].Value;
                if (part != null)
                    sb.Append(part);
            }
            return double.Parse(sb.ToString(), CultureInfo.InvariantCulture);
        }
    }
}