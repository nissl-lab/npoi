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
    using TestCases.SS;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.HSSF.Util;
    using NPOI.HSSF.UserModel;
    using System.Text;
    using NPOI.SS;

    /**
     * Common superclass for testing implementatiosn of
     *  {@link NPOI.SS.usermodel.Cell}
     */
    public class BaseTestCell
    {

        protected ITestDataProvider _testDataProvider;

        public BaseTestCell()
            : this(TestCases.HSSF.HSSFITestDataProvider.Instance)
        { }

        /**
         * @param testDataProvider an object that provides test data in HSSF / XSSF specific way
         */
        protected BaseTestCell(ITestDataProvider testDataProvider)
        {
            // One or more test methods depends on the american culture.
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            _testDataProvider = testDataProvider;
        }
        [Test]
        public void TestSetValues()
        {
            IWorkbook book = _testDataProvider.CreateWorkbook();
            ISheet sheet = book.CreateSheet("test");
            IRow row = sheet.CreateRow(0);

            ICreationHelper factory = book.GetCreationHelper();
            ICell cell = row.CreateCell(0);

            cell.SetCellValue(1.2);
            Assert.AreEqual(1.2, cell.NumericCellValue, 0.0001);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            AssertProhibitedValueAccess(cell, CellType.Boolean, CellType.String,
                    CellType.Formula, CellType.Error);

            cell.SetCellValue(false);
            Assert.AreEqual(false, cell.BooleanCellValue);
            Assert.AreEqual(CellType.Boolean, cell.CellType);
            cell.SetCellValue(true);
            Assert.AreEqual(true, cell.BooleanCellValue);
            AssertProhibitedValueAccess(cell, CellType.Numeric, CellType.String,
                    CellType.Formula, CellType.Error);

            cell.SetCellValue(factory.CreateRichTextString("Foo"));
            Assert.AreEqual("Foo", cell.RichStringCellValue.String);
            Assert.AreEqual("Foo", cell.StringCellValue);
            Assert.AreEqual(CellType.String, cell.CellType);
            AssertProhibitedValueAccess(cell, CellType.Numeric, CellType.Boolean,
                    CellType.Formula, CellType.Error);

            cell.SetCellValue("345");
            Assert.AreEqual("345", cell.RichStringCellValue.String);
            Assert.AreEqual("345", cell.StringCellValue);
            Assert.AreEqual(CellType.String, cell.CellType);
            AssertProhibitedValueAccess(cell, CellType.Numeric, CellType.Boolean,
                    CellType.Formula, CellType.Error);

            DateTime dt = DateTime.Now.AddMilliseconds(123456789);
            cell.SetCellValue(dt);
            Assert.IsTrue((dt.Ticks - cell.DateCellValue.Ticks) >= -20000);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            AssertProhibitedValueAccess(cell, CellType.Boolean, CellType.String,
                    CellType.Formula, CellType.Error);

            cell.SetCellValue(dt);
            Assert.IsTrue((dt.Ticks - cell.DateCellValue.Ticks) >= -20000);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            AssertProhibitedValueAccess(cell, CellType.Boolean, CellType.String,
                    CellType.Formula, CellType.Error);

            cell.SetCellErrorValue(FormulaError.NA.Code);
            Assert.AreEqual(FormulaError.NA.Code, cell.ErrorCellValue);
            Assert.AreEqual(CellType.Error, cell.CellType);
            AssertProhibitedValueAccess(cell, CellType.Numeric, CellType.Boolean,
                    CellType.Formula, CellType.String);
        }

        private static void AssertProhibitedValueAccess(ICell cell, params CellType[] types)
        {
            object a;
            foreach (CellType type in types)
            {
                try
                {
                    switch (type)
                    {
                        case CellType.Numeric:
                            a = cell.NumericCellValue;
                            break;
                        case CellType.String:
                            a = cell.StringCellValue;
                            break;
                        case CellType.Boolean:
                            a = cell.BooleanCellValue;
                            break;
                        case CellType.Formula:
                            a = cell.CellFormula;
                            break;
                        case CellType.Error:
                            a = cell.ErrorCellValue;
                            break;
                    }
                    Assert.Fail("Should get exception when Reading cell type (" + type + ").");
                }
                catch (InvalidOperationException e)
                {
                    // expected during successful test
                    Assert.IsTrue(e.Message.StartsWith("Cannot get a"));
                }
            }
        }

        /**
         * test that Boolean and Error types (BoolErrRecord) are supported properly.
         */
        [Test]
        public void TestBoolErr()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet("testSheet1");
            IRow r;
            ICell c;
            r = s.CreateRow(0);
            c = r.CreateCell(1);
            //c.SetCellType(HSSFCellType.Boolean);
            c.SetCellValue(true);

            c = r.CreateCell(2);
            //c.SetCellType(HSSFCellType.Boolean);
            c.SetCellValue(false);

            r = s.CreateRow(1);
            c = r.CreateCell(1);
            //c.SetCellType(HSSFCellType.Error);
            c.SetCellErrorValue((byte)0);

            c = r.CreateCell(2);
            //c.SetCellType(HSSFCellType.Error);
            c.SetCellErrorValue((byte)7);

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(1);
            Assert.IsTrue(c.BooleanCellValue, "bool value 0,1 = true");
            c = r.GetCell(2);
            Assert.IsTrue(c.BooleanCellValue == false, "bool value 0,2 = false");
            r = s.GetRow(1);
            c = r.GetCell(1);
            Assert.IsTrue(c.ErrorCellValue == 0, "bool value 0,1 = 0");
            c = r.GetCell(2);
            Assert.IsTrue(c.ErrorCellValue == 7, "bool value 0,2 = 7");
        }

        /**
         * test that Cell Styles being applied to formulas remain intact
         */
        [Test]
        public void TestFormulaStyle()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet("testSheet1");
            IRow r = null;
            ICell c = null;
            ICellStyle cs = wb.CreateCellStyle();
            IFont f = wb.CreateFont();
            f.FontHeightInPoints = 20;
            f.Color = (HSSFColor.Red.Index);
            f.Boldweight = (int)FontBoldWeight.Bold;
            f.FontName = "Arial Unicode MS";
            cs.FillBackgroundColor = 3;
            cs.SetFont(f);
            cs.BorderTop = BorderStyle.Thin;
            cs.BorderRight = BorderStyle.Thin;
            cs.BorderLeft = BorderStyle.Thin;
            cs.BorderBottom = BorderStyle.Thin;

            r = s.CreateRow(0);
            c = r.CreateCell(0);
            c.CellStyle = cs;
            c.CellFormula = ("2*3");

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            c = r.GetCell(0);

            Assert.IsTrue((c.CellType == CellType.Formula), "Formula Cell at 0,0");
            cs = c.CellStyle;

            Assert.IsNotNull(cs, "Formula Cell Style");
            Assert.AreEqual(f.Index, cs.FontIndex, "Font Index Matches");
            Assert.AreEqual(BorderStyle.Thin, cs.BorderTop , "Top Border");
            Assert.AreEqual(BorderStyle.Thin, cs.BorderLeft, "Left Border");
            Assert.AreEqual(BorderStyle.Thin, cs.BorderRight, "Right Border");
            Assert.AreEqual(BorderStyle.Thin, cs.BorderBottom, "Bottom Border");
        }

        /**tests the ToString() method of HSSFCell*/
        [Test]
        public void TestToString()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IRow r = wb.CreateSheet("Sheet1").CreateRow(0);
            ICreationHelper factory = wb.GetCreationHelper();

            r.CreateCell(0).SetCellValue(true);
            r.CreateCell(1).SetCellValue(1.5);
            r.CreateCell(2).SetCellValue(factory.CreateRichTextString("Astring"));
            r.CreateCell(3).SetCellErrorValue((byte)ErrorConstants.ERROR_DIV_0);
            r.CreateCell(4).CellFormula = ("A1+B1");

            Assert.AreEqual("TRUE", r.GetCell(0).ToString(), "Boolean");
            Assert.AreEqual("1.5", r.GetCell(1).ToString(), "Numeric");
            Assert.AreEqual("Astring", r.GetCell(2).ToString(), "String");
            Assert.AreEqual("#DIV/0!", r.GetCell(3).ToString(), "Error");
            Assert.AreEqual("A1+B1", r.GetCell(4).ToString(), "Formula");

            //Write out the file, read it in, and then check cell values
            wb = _testDataProvider.WriteOutAndReadBack(wb);

            r = wb.GetSheetAt(0).GetRow(0);
            Assert.AreEqual("TRUE", r.GetCell(0).ToString(), "Boolean");
            Assert.AreEqual("1.5", r.GetCell(1).ToString(), "Numeric");
            Assert.AreEqual("Astring", r.GetCell(2).ToString(), "String");
            Assert.AreEqual("#DIV/0!", r.GetCell(3).ToString(), "Error");
            Assert.AreEqual("A1+B1", r.GetCell(4).ToString(), "Formula");
        }

        /**
         *  Test that Setting cached formula result keeps the cell type
         */
        [Test]
        public void TestSetFormulaValue()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);

            ICell c1 = r.CreateCell(0);
            c1.CellFormula = ("NA()");
            Assert.AreEqual(0.0, c1.NumericCellValue, 0.0);
            Assert.AreEqual(CellType.Numeric, c1.CachedFormulaResultType);
            c1.SetCellValue(10);
            Assert.AreEqual(10.0, c1.NumericCellValue, 0.0);
            Assert.AreEqual(CellType.Formula, c1.CellType);
            Assert.AreEqual(CellType.Numeric, c1.CachedFormulaResultType);

            ICell c2 = r.CreateCell(1);
            c2.CellFormula = ("NA()");
            Assert.AreEqual(0.0, c2.NumericCellValue, 0.0);
            Assert.AreEqual(CellType.Numeric, c2.CachedFormulaResultType);
            c2.SetCellValue("I Changed!");
            Assert.AreEqual("I Changed!", c2.StringCellValue);
            Assert.AreEqual(CellType.Formula, c2.CellType);
            Assert.AreEqual(CellType.String, c2.CachedFormulaResultType);

            //calglin Cell.CellFormula = (null) for a non-formula cell
            ICell c3 = r.CreateCell(2);
            c3.CellFormula = (null);
            Assert.AreEqual(CellType.Blank, c3.CellType);

        }
        private ICell CreateACell()
        {
            return _testDataProvider.CreateWorkbook().CreateSheet("Sheet1").CreateRow(0).CreateCell(0);
        }

        [Test]
        public void TestChangeTypeStringToBool()
        {
            ICell cell = CreateACell();

            cell.SetCellValue("TRUE");
            Assert.AreEqual(CellType.String, cell.CellType);
            try
            {
                cell.SetCellType(CellType.Boolean);
            }
            catch (InvalidCastException)
            {
                throw new AssertionException(
                        "Identified bug in conversion of cell from text to bool");
            }

            Assert.AreEqual(CellType.Boolean, cell.CellType);
            Assert.AreEqual(true, cell.BooleanCellValue);
            cell.SetCellType(CellType.String);
            Assert.AreEqual("TRUE", cell.RichStringCellValue.String);

            // 'false' text to bool and back
            cell.SetCellValue("FALSE");
            cell.SetCellType(CellType.Boolean);
            Assert.AreEqual(CellType.Boolean, cell.CellType);
            Assert.AreEqual(false, cell.BooleanCellValue);
            cell.SetCellType(CellType.String);
            Assert.AreEqual("FALSE", cell.RichStringCellValue.String);
        }
        [Test]
        public void TestChangeTypeBoolToString()
        {
            ICell cell = CreateACell();

            cell.SetCellValue(true);
            try
            {
                cell.SetCellType(CellType.String);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.Equals("Cannot get a text value from a bool cell"))
                {
                    throw new AssertionException(
                            "Identified bug in conversion of cell from bool to text");
                }
                throw e;
            }
            Assert.AreEqual("TRUE", cell.RichStringCellValue.String);
        }
        [Test]
        public void TestChangeTypeErrorToNumber()
        {
            ICell cell = CreateACell();
            cell.SetCellErrorValue((byte)ErrorConstants.ERROR_NAME);
            try
            {
                cell.SetCellValue(2.5);
            }
            catch (InvalidCastException)
            {
                throw new AssertionException("Identified bug 46479b");
            }
            Assert.AreEqual(2.5, cell.NumericCellValue, 0.0);
        }
        [Test]
        public void TestChangeTypeErrorToBoolean()
        {
            ICell cell = CreateACell();
            cell.SetCellErrorValue((byte)ErrorConstants.ERROR_NAME);
            cell.SetCellValue(true);
            try
            {
                object a = cell.BooleanCellValue;
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.Equals("Cannot get a bool value from a error cell"))
                {

                    throw new AssertionException("Identified bug 46479c");
                }
                throw e;
            }
            Assert.AreEqual(true, cell.BooleanCellValue);
        }

        /**
         * Test for a bug observed around svn r886733 when using
         * {@link FormulaEvaluator#EvaluateInCell(Cell)} with a
         * string result type.
         */
        [Test]
        public void TestConvertStringFormulaCell()
        {
            ICell cellA1 = CreateACell();
            cellA1.CellFormula = ("\"abc\"");

            // default cached formula result is numeric zero
            Assert.AreEqual(0.0, cellA1.NumericCellValue, 0.0);

            IFormulaEvaluator fe = cellA1.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();

            fe.EvaluateFormulaCell(cellA1);
            Assert.AreEqual("abc", cellA1.StringCellValue);

            fe.EvaluateInCell(cellA1);
            if (cellA1.StringCellValue.Equals(""))
            {
                throw new AssertionException("Identified bug with writing back formula result of type string");
            }
            Assert.AreEqual("abc", cellA1.StringCellValue);
        }
        /**
         * similar to {@link #testConvertStringFormulaCell()} but  Checks at a
         * lower level that {#link {@link Cell#SetCellType(int)} works properly
         */
        [Test]
        public void TestSetTypeStringOnFormulaCell()
        {
            ICell cellA1 = CreateACell();
            IFormulaEvaluator fe = cellA1.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();

            cellA1.CellFormula = ("\"DEF\"");
            fe.ClearAllCachedResultValues();
            fe.EvaluateFormulaCell(cellA1);
            Assert.AreEqual("DEF", cellA1.StringCellValue);
            cellA1.SetCellType(CellType.String);
            Assert.AreEqual("DEF", cellA1.StringCellValue);

            cellA1.CellFormula = ("25.061");
            fe.ClearAllCachedResultValues();
            fe.EvaluateFormulaCell(cellA1);
            ConfirmCannotReadString(cellA1);
            Assert.AreEqual(25.061, cellA1.NumericCellValue, 0.0);
            cellA1.SetCellType(CellType.String);
            Assert.AreEqual("25.061", cellA1.StringCellValue);

            cellA1.CellFormula = ("TRUE");
            fe.ClearAllCachedResultValues();
            fe.EvaluateFormulaCell(cellA1);
            ConfirmCannotReadString(cellA1);
            Assert.AreEqual(true, cellA1.BooleanCellValue);
            cellA1.SetCellType(CellType.String);
            Assert.AreEqual("TRUE", cellA1.StringCellValue);

            cellA1.CellFormula = ("#NAME?");
            fe.ClearAllCachedResultValues();
            fe.EvaluateFormulaCell(cellA1);
            ConfirmCannotReadString(cellA1);
            Assert.AreEqual(ErrorConstants.ERROR_NAME, cellA1.ErrorCellValue);
            cellA1.SetCellType(CellType.String);
            Assert.AreEqual("#NAME?", cellA1.StringCellValue);
        }

        private static void ConfirmCannotReadString(ICell cell)
        {
            AssertProhibitedValueAccess(cell, CellType.String);
        }

        /**
         * Test for bug in ConvertCellValueToBoolean to make sure that formula results get Converted
         */
        [Test]
        public void TestChangeTypeFormulaToBoolean()
        {
            ICell cell = CreateACell();
            cell.CellFormula = ("1=1");
            cell.SetCellValue(true);
            cell.SetCellType(CellType.Boolean);
            if (cell.BooleanCellValue == false)
            {
                throw new AssertionException("Identified bug 46479d");
            }
            Assert.AreEqual(true, cell.BooleanCellValue);
        }

        /**
         * Bug 40296:	  HSSFCell.CellFormula =  throws
         *   InvalidCastException if cell is Created using HSSFRow.CreateCell(short column, int type)
         */
        [Test]
        public void Test40296()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet workSheet = wb.CreateSheet("Sheet1");
            ICell cell;
            IRow row = workSheet.CreateRow(0);

            cell = row.CreateCell(0, CellType.Numeric);
            cell.SetCellValue(1.0);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            Assert.AreEqual(1.0, cell.NumericCellValue, 0.0);

            cell = row.CreateCell(1, CellType.Numeric);
            cell.SetCellValue(2.0);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            Assert.AreEqual(2.0, cell.NumericCellValue, 0.0);

            cell = row.CreateCell(2, CellType.Formula);
            cell.CellFormula = ("SUM(A1:B1)");
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual("SUM(A1:B1)", cell.CellFormula);

            //serialize and check again
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            row = wb.GetSheetAt(0).GetRow(0);
            cell = row.GetCell(0);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            Assert.AreEqual(1.0, cell.NumericCellValue, 0.0);

            cell = row.GetCell(1);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            Assert.AreEqual(2.0, cell.NumericCellValue, 0.0);

            cell = row.GetCell(2);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual("SUM(A1:B1)", cell.CellFormula);
        }
        [Test]
        public void TestSetStringInFormulaCell_bug44606()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICell cell = wb.CreateSheet("Sheet1").CreateRow(0).CreateCell(0);
            cell.CellFormula = ("B1&C1");
            try
            {
                cell.SetCellValue(wb.GetCreationHelper().CreateRichTextString("hello"));
            }
            catch (InvalidCastException)
            {
                throw new AssertionException("Identified bug 44606");
            }
        }

        /**
         *  Make sure that cell.SetCellType(Cell.CELL_TYPE_BLANK) preserves the cell style
         */
        [Test]
        public void TestSetBlank_bug47028()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICellStyle style = wb.CreateCellStyle();
            ICell cell = wb.CreateSheet("Sheet1").CreateRow(0).CreateCell(0);
            cell.CellStyle = (style);
            int i1 = cell.CellStyle.Index;
            cell.SetCellType(CellType.Blank);
            int i2 = cell.CellStyle.Index;
            Assert.AreEqual(i1, i2);
        }

        /**
         * Excel's implementation of floating number arithmetic does not fully adhere to IEEE 754:
         *
         * From http://support.microsoft.com/kb/78113:
         *
         * <ul>
         * <li> Positive/Negative InfInities:
         *   InfInities occur when you divide by 0. Excel does not support infInities, rather,
         *   it gives a #DIV/0! error in these cases.
         * </li>
         * <li>
         *   Not-a-Number (NaN):
         *   NaN is used to represent invalid operations (such as infInity/infinity, 
         *   infInity-infinity, or the square root of -1). NaNs allow a program to
         *   continue past an invalid operation. Excel instead immediately generates
         *   an error such as #NUM! or #DIV/0!.
         * </li>
         * </ul>
         */
        [Test]
        public void TestNanAndInfInity()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet workSheet = wb.CreateSheet("Sheet1");
            IRow row = workSheet.CreateRow(0);

            ICell cell0 = row.CreateCell(0);
            cell0.SetCellValue(Double.NaN);
            Assert.AreEqual(CellType.Error, cell0.CellType, "Double.NaN should change cell type to CELL_TYPE_ERROR");
            Assert.AreEqual(ErrorConstants.ERROR_NUM, cell0.ErrorCellValue, "Double.NaN should change cell value to #NUM!");

            ICell cell1 = row.CreateCell(1);
            cell1.SetCellValue(Double.PositiveInfinity);
            Assert.AreEqual(CellType.Error, cell1.CellType, "Double.PositiveInfinity should change cell type to CELL_TYPE_ERROR");
            Assert.AreEqual(ErrorConstants.ERROR_DIV_0, cell1.ErrorCellValue, "Double.POSITIVE_INFINITY should change cell value to #DIV/0!");

            ICell cell2 = row.CreateCell(2);
            cell2.SetCellValue(Double.NegativeInfinity);
            Assert.AreEqual(CellType.Error, cell2.CellType, "Double.NegativeInfinity should change cell type to CELL_TYPE_ERROR");
            Assert.AreEqual(ErrorConstants.ERROR_DIV_0, cell2.ErrorCellValue, "Double.NEGATIVE_INFINITY should change cell value to #DIV/0!");

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            row = wb.GetSheetAt(0).GetRow(0);

            cell0 = row.GetCell(0);
            Assert.AreEqual(CellType.Error, cell0.CellType);
            Assert.AreEqual(ErrorConstants.ERROR_NUM, cell0.ErrorCellValue);

            cell1 = row.GetCell(1);
            Assert.AreEqual(CellType.Error, cell1.CellType);
            Assert.AreEqual(ErrorConstants.ERROR_DIV_0, cell1.ErrorCellValue);

            cell2 = row.GetCell(2);
            Assert.AreEqual(CellType.Error, cell2.CellType);
            Assert.AreEqual(ErrorConstants.ERROR_DIV_0, cell2.ErrorCellValue);
        }
        [Test]
        public void TestDefaultStyleProperties()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();

            ICell cell = wb.CreateSheet("Sheet1").CreateRow(0).CreateCell(0);
            ICellStyle style = cell.CellStyle;

            Assert.IsTrue(style.IsLocked);
            Assert.IsFalse(style.IsHidden);
            Assert.AreEqual(0, style.Indention);
            Assert.AreEqual(0, style.FontIndex);
            Assert.AreEqual(0, (int)style.Alignment);
            Assert.AreEqual(0, style.DataFormat);
            Assert.AreEqual(false, style.WrapText);

            ICellStyle style2 = wb.CreateCellStyle();
            Assert.IsTrue(style2.IsLocked);
            Assert.IsFalse(style2.IsHidden);
            style2.IsLocked = (/*setter*/false);
            style2.IsHidden = (/*setter*/true);
            Assert.IsFalse(style2.IsLocked);
            Assert.IsTrue(style2.IsHidden);

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            cell = wb.GetSheetAt(0).GetRow(0).GetCell(0);
            style = cell.CellStyle;
            Assert.IsFalse(style2.IsLocked);
            Assert.IsTrue(style2.IsHidden);

            style2.IsLocked = (/*setter*/true);
            style2.IsHidden = (/*setter*/false);
            Assert.IsTrue(style2.IsLocked);
            Assert.IsFalse(style2.IsHidden);
        }
        [Test]
        public void TestBug55658SetNumericValue()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row = sh.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue(23);

            cell.SetCellValue("some");

            cell = row.CreateCell(1);
            cell.SetCellValue(23);

            cell.SetCellValue("24");

            wb = _testDataProvider.WriteOutAndReadBack(wb);

            Assert.AreEqual("some", wb.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual("24", wb.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue);
        }
        [Test]
        public void TestRemoveHyperlink()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet("test");
            IRow row = sh.CreateRow(0);
            ICreationHelper helper = wb.GetCreationHelper();

            ICell cell1 = row.CreateCell(1);
            IHyperlink link1 = helper.CreateHyperlink(HyperlinkType.Url);
            cell1.Hyperlink = (/*setter*/link1);
            Assert.IsNotNull(cell1.Hyperlink);
            cell1.RemoveHyperlink();
            Assert.IsNull(cell1.Hyperlink);

            ICell cell2 = row.CreateCell(0);
            IHyperlink link2 = helper.CreateHyperlink(HyperlinkType.Url);
            cell2.Hyperlink = (/*setter*/link2);
            Assert.IsNotNull(cell2.Hyperlink);
            cell2.Hyperlink = (/*setter*/null);
            Assert.IsNull(cell2.Hyperlink);

            ICell cell3 = row.CreateCell(2);
            IHyperlink link3 = helper.CreateHyperlink(HyperlinkType.Url);
            link3.Address = (/*setter*/"http://poi.apache.org/");
            cell3.Hyperlink = (/*setter*/link3);
            Assert.IsNotNull(cell3.Hyperlink);

            IWorkbook wbBack = _testDataProvider.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);

            cell1 = wbBack.GetSheet("test").GetRow(0).GetCell(1);
            Assert.IsNull(cell1.Hyperlink);
            cell2 = wbBack.GetSheet("test").GetRow(0).GetCell(0);
            Assert.IsNull(cell2.Hyperlink);
            cell3 = wbBack.GetSheet("test").GetRow(0).GetCell(2);
            Assert.IsNotNull(cell3.Hyperlink);
        }

        /**
         * Cell with the formula that returns error must return error code(There was
         * an problem that cell could not return error value form formula cell).
         * @
         */
        [Test]
        public void TestGetErrorCellValueFromFormulaCell()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);
                cell.CellFormula = (/*setter*/"SQRT(-1)");
                wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateFormulaCell(cell);
                Assert.AreEqual(36, cell.ErrorCellValue);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestSetRemoveStyle()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            // different default style indexes for HSSF and XSSF/SXSSF
            ICellStyle defaultStyle = wb.GetCellStyleAt(wb is HSSFWorkbook ? (short)15 : (short)0);

            // Starts out with the default style
            Assert.AreEqual(defaultStyle, cell.CellStyle);

            // Create some styles, no change
            ICellStyle style1 = wb.CreateCellStyle();
            ICellStyle style2 = wb.CreateCellStyle();
            style1.DataFormat = (/*setter*/(short)2);
            style2.DataFormat = (/*setter*/(short)3);

            Assert.AreEqual(defaultStyle, cell.CellStyle);

            // Apply one, Changes
            cell.CellStyle = (/*setter*/style1);
            Assert.AreEqual(style1, cell.CellStyle);

            // Apply the other, Changes
            cell.CellStyle = (/*setter*/style2);
            Assert.AreEqual(style2, cell.CellStyle);

            // Remove, goes back to default
            cell.CellStyle = (/*setter*/null);
            Assert.AreEqual(defaultStyle, cell.CellStyle);

            // Add back, returns
            cell.CellStyle = (/*setter*/style2);
            Assert.AreEqual(style2, cell.CellStyle);

            wb.Close();
        }

        [Test]
        public void Test57008()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            IRow row0 = sheet.CreateRow(0);
            ICell cell0 = row0.CreateCell(0);
            cell0.SetCellValue("row 0, cell 0 _x0046_ without Changes");

            ICell cell1 = row0.CreateCell(1);
            cell1.SetCellValue("row 0, cell 1 _x005fx0046_ with Changes");

            ICell cell2 = row0.CreateCell(2);
            cell2.SetCellValue("hgh_x0041_**_x0100_*_x0101_*_x0190_*_x0200_*_x0300_*_x0427_*");

            CheckUnicodeValues(wb);

            //		String fname = "/tmp/Test_xNNNN_inCell" + (wb is HSSFWorkbook ? ".xls" : ".xlsx");
            //		FileOutputStream out1 = new FileOutputStream(fname);
            //		try {
            //			wb.Write(out1);
            //		} finally {
            //			out1.Close();
            //		}

            IWorkbook wbBack = _testDataProvider.WriteOutAndReadBack(wb);
            CheckUnicodeValues(wbBack);
        }

        protected void CheckUnicodeValues(IWorkbook wb)
        {
            Assert.AreEqual((wb is HSSFWorkbook ? "row 0, cell 0 _x0046_ without Changes" : "row 0, cell 0 F without Changes"),
                    wb.GetSheetAt(0).GetRow(0).GetCell(0).ToString());
            Assert.AreEqual((wb is HSSFWorkbook ? "row 0, cell 1 _x005fx0046_ with Changes" : "row 0, cell 1 _x005fx0046_ with Changes"),
                    wb.GetSheetAt(0).GetRow(0).GetCell(1).ToString());
            Assert.AreEqual((wb is HSSFWorkbook ? "hgh_x0041_**_x0100_*_x0101_*_x0190_*_x0200_*_x0300_*_x0427_*" : "hghA**\u0100*\u0101*\u0190*\u0200*\u0300*\u0427*"),
                    wb.GetSheetAt(0).GetRow(0).GetCell(2).ToString());
        }

        /**
         *  The maximum length of cell contents (text) is 32,767 characters.
         * @
         */
        [Test]
        public void TestMaxTextLength()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            ICell cell = sheet.CreateRow(0).CreateCell(0);

            int maxlen = wb is HSSFWorkbook ?
                    SpreadsheetVersion.EXCEL97.MaxTextLength
                    : SpreadsheetVersion.EXCEL2007.MaxTextLength;
            Assert.AreEqual(32767, maxlen);

            StringBuilder b = new StringBuilder();

            // 32767 is okay
            for (int i = 0; i < maxlen; i++)
            {
                b.Append("X");
            }
            cell.SetCellValue(b.ToString());

            b.Append("X");
            // 32768 produces an invalid XLS file
            try
            {
                cell.SetCellValue(b.ToString());
                Assert.Fail("Expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("The maximum length of cell contents (text) is 32,767 characters", e.Message);
            }
            wb.Close();
        }

    }
}