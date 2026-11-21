/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace TestCases.XSSF.UserModel
{
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Test built-in table styles
    /// </summary>
    public class TestTableStyles
    {

        /// <summary>
        /// Test that a built-in style is initialized properly
        /// </summary>
        [Test]
        public void TestBuiltinStyleInit()
        {
            ITableStyle style = XSSFBuiltinTableStyle.GetStyle(XSSFBuiltinTableStyleEnum.TableStyleMedium2);
            ClassicAssert.IsNotNull(style, "no style found for Medium2");
            ClassicAssert.IsNull(style.GetStyle(TableStyleType.blankRow), "Should not have style info for blankRow");
            IDifferentialStyleProvider headerRow = style.GetStyle(TableStyleType.headerRow);
            ClassicAssert.IsNotNull(headerRow, "no header row style");
            IFontFormatting font = headerRow.FontFormatting;
            ClassicAssert.IsNotNull(font, "No header row font formatting");
            ClassicAssert.IsTrue(font.IsBold, "header row not bold");
            IPatternFormatting fill = headerRow.PatternFormatting;
            ClassicAssert.IsNotNull(fill, "No header fill");
            ClassicAssert.AreEqual(4, ((XSSFColor)fill.FillBackgroundColorColor).Theme, "wrong header fill");
        }

        [Test]
        public void TestCustomStyle()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("tableStyle.xlsx");

            ITable table = wb.GetTable("Table1");
            ClassicAssert.IsNotNull(table, "missing table");

            ITableStyleInfo style = table.Style;
            ClassicAssert.IsNotNull(style, "Missing table style info");
            ClassicAssert.IsNotNull(style.Style, "Missing table style");
            ClassicAssert.AreEqual("TestTableStyle", style.Name, "Wrong name");
            ClassicAssert.AreEqual("TestTableStyle", style.Style.Name, "Wrong name");

            IDifferentialStyleProvider firstColumn = style.Style.GetStyle(TableStyleType.firstColumn);
            ClassicAssert.IsNotNull(firstColumn, "no first column style");
            IFontFormatting font = firstColumn.FontFormatting;
            ClassicAssert.IsNotNull(font, "no first col font");
            ClassicAssert.IsTrue(font.IsBold, "wrong first col bold");

            wb.Close();
        }
    }
}
