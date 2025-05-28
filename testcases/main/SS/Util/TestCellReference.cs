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

using NPOI.SS.Util;
using NUnit.Framework;using NUnit.Framework.Legacy;
using TestCases.HSSF.Record;

using System.IO;
using System;
using NPOI.SS;

namespace TestCases.SS.Util
{
    /**
     * Tests that the common CellReference works as we need it to
     */
    [TestFixture]
    public class TestCellReference
    {
        [Test]
        public void TestConstructors()
        {
            CellReference cellReference;
            String sheet = "Sheet1";
            String cellRef = "A1";
            int row = 0;
            int col = 0;
            bool absRow = true;
            bool absCol = false;

            cellReference = new CellReference(row, col);
            ClassicAssert.AreEqual("A1", cellReference.FormatAsString());

            cellReference = new CellReference(row, col, absRow, absCol);
            ClassicAssert.AreEqual("A$1", cellReference.FormatAsString());

            cellReference = new CellReference(row, (short)col);
            ClassicAssert.AreEqual("A1", cellReference.FormatAsString());

            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual("A1", cellReference.FormatAsString());

            cellReference = new CellReference(sheet, row, col, absRow, absCol);
            ClassicAssert.AreEqual("Sheet1!A$1", cellReference.FormatAsString());
        }

        [Test]
        public void TestFormatAsString()
        {
            CellReference cellReference;

            cellReference = new CellReference(null, 0, 0, false, false);
            ClassicAssert.AreEqual("A1", cellReference.FormatAsString());

            //absolute references
            cellReference = new CellReference(null, 0, 0, true, false);
            ClassicAssert.AreEqual("A$1", cellReference.FormatAsString());

            //sheet name with no spaces
            cellReference = new CellReference("Sheet1", 0, 0, true, false);
            ClassicAssert.AreEqual("Sheet1!A$1", cellReference.FormatAsString());

            //sheet name with spaces
            cellReference = new CellReference("Sheet 1", 0, 0, true, false);
            ClassicAssert.AreEqual("'Sheet 1'!A$1", cellReference.FormatAsString());
        }


        [Test]
        public void TestGetCellRefParts()
        {
            CellReference cellReference;
            String[] parts;

            String cellRef = "A1";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(0, cellReference.Col);
            parts = cellReference.CellRefParts;
            ClassicAssert.IsNotNull(parts);
            ClassicAssert.AreEqual(null, parts[0]);
            ClassicAssert.AreEqual("1", parts[1]);
            ClassicAssert.AreEqual("A", parts[2]);

            cellRef = "AA1";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26, cellReference.Col);
            parts = cellReference.CellRefParts;
            ClassicAssert.IsNotNull(parts);
            ClassicAssert.AreEqual(null, parts[0]);
            ClassicAssert.AreEqual("1", parts[1]);
            ClassicAssert.AreEqual("AA", parts[2]);

            cellRef = "AA100";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26, cellReference.Col);
            parts = cellReference.CellRefParts;
            ClassicAssert.IsNotNull(parts);
            ClassicAssert.AreEqual(null, parts[0]);
            ClassicAssert.AreEqual("100", parts[1]);
            ClassicAssert.AreEqual("AA", parts[2]);

            cellRef = "AAA300";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(702, cellReference.Col);
            parts = cellReference.CellRefParts;
            ClassicAssert.IsNotNull(parts);
            ClassicAssert.AreEqual(null, parts[0]);
            ClassicAssert.AreEqual("300", parts[1]);
            ClassicAssert.AreEqual("AAA", parts[2]);

            cellRef = "ZZ100521";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26 * 26 + 25, cellReference.Col);
            parts = cellReference.CellRefParts;
            ClassicAssert.IsNotNull(parts);
            ClassicAssert.AreEqual(null, parts[0]);
            ClassicAssert.AreEqual("100521", parts[1]);
            ClassicAssert.AreEqual("ZZ", parts[2]);

            cellRef = "ZYX987";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26 * 26 * 26 + 25 * 26 + 24 - 1, cellReference.Col);
            parts = cellReference.CellRefParts;
            ClassicAssert.IsNotNull(parts);
            ClassicAssert.AreEqual(null, parts[0]);
            ClassicAssert.AreEqual("987", parts[1]);
            ClassicAssert.AreEqual("ZYX", parts[2]);

            cellRef = "AABC10065";
            cellReference = new CellReference(cellRef);
            parts = cellReference.CellRefParts;
            ClassicAssert.IsNotNull(parts);
            ClassicAssert.AreEqual(null, parts[0]);
            ClassicAssert.AreEqual("10065", parts[1]);
            ClassicAssert.AreEqual("AABC", parts[2]);
        }
        [Test]
        public void TestGetColNumFromRef()
        {
            String cellRef = "A1";
            CellReference cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(0, cellReference.Col);

            cellRef = "AA1";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26, cellReference.Col);

            cellRef = "AB1";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(27, cellReference.Col);

            cellRef = "BA1";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26 + 26, cellReference.Col);

            cellRef = "CA1";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26 + 26 + 26, cellReference.Col);

            cellRef = "ZA1";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26 * 26, cellReference.Col);

            cellRef = "ZZ1";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26 * 26 + 25, cellReference.Col);

            cellRef = "AAA1";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(26 * 26 + 26, cellReference.Col);


            cellRef = "A1100";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(0, cellReference.Col);

            cellRef = "BC15";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(54, cellReference.Col);
        }
        [Test]
        public void TestGetRowNumFromRef()
        {
            String cellRef = "A1";
            CellReference cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(0, cellReference.Row);

            cellRef = "A12";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(11, cellReference.Row);

            cellRef = "AS121";
            cellReference = new CellReference(cellRef);
            ClassicAssert.AreEqual(120, cellReference.Row);
        }
        [Test]
        public void TestConvertNumToColString()
        {
            short col = 702;
            String collRef = new CellReference(0, col).FormatAsString();
            ClassicAssert.AreEqual("AAA1", collRef);

            short col2 = 0;
            String collRef2 = new CellReference(0, col2).FormatAsString();
            ClassicAssert.AreEqual("A1", collRef2);

            short col3 = 27;
            String collRef3 = new CellReference(0, col3).FormatAsString();
            ClassicAssert.AreEqual("AB1", collRef3);

            short col4 = 2080;
            String collRef4 = new CellReference(0, col4).FormatAsString();
            ClassicAssert.AreEqual("CBA1", collRef4);
        }
        [Test]
        public void TestBadRowNumber()
        {
            SpreadsheetVersion v97 = SpreadsheetVersion.EXCEL97;
            SpreadsheetVersion v2007 = SpreadsheetVersion.EXCEL2007;

            ConfirmCrInRange(true, "A", "1", v97);
            ConfirmCrInRange(true, "IV", "65536", v97);
            ConfirmCrInRange(false, "IV", "65537", v97);
            ConfirmCrInRange(false, "IW", "65536", v97);

            ConfirmCrInRange(true, "A", "1", v2007);
            ConfirmCrInRange(true, "XFD", "1048576", v2007);
            ConfirmCrInRange(false, "XFD", "1048577", v2007);
            ConfirmCrInRange(false, "XFE", "1048576", v2007);

            if (CellReference.CellReferenceIsWithinRange("B", "0", v97))
            {
                throw new AssertionException ("Identified bug 47312a");
            }

            ConfirmCrInRange(false, "A", "0", v97);
            ConfirmCrInRange(false, "A", "0", v2007);
        }
        [Test]
        public void TestInvalidReference()
        {
            try
            {
                new CellReference("Sheet1!#REF!");
                Assert.Fail("Shouldn't be able to create a #REF! refence");
            }
            catch (ArgumentException) { }

            try
            {
                new CellReference("'MySheetName'!#REF!");
                Assert.Fail("Shouldn't be able to create a #REF! refence");
            }
            catch (ArgumentException) { }

            try
            {
                new CellReference("#REF!");
                Assert.Fail("Shouldn't be able to create a #REF! refence");
            }
            catch (ArgumentException) { }
        }

        private static void ConfirmCrInRange(bool expResult, String colStr, String rowStr,
                SpreadsheetVersion sv)
        {
            if (expResult == CellReference.CellReferenceIsWithinRange(colStr, rowStr, sv))
            {
                return;
            }
            throw new AssertionException("expected (c='" + colStr + "', r='" + rowStr + "' to be "
                    + (expResult ? "within" : "out of") + " bounds for version " + sv.ToString());
        }

        /**
         * bug 59684: separateRefParts Assert.Fails on entire-column references
         */
        [Test]
        public void EntireColumnReferences()
        {
            CellReference ref1 = new CellReference("HOME!$169");
            ClassicAssert.AreEqual("HOME", ref1.SheetName);
            ClassicAssert.AreEqual(168, ref1.Row);
            ClassicAssert.AreEqual(-1, ref1.Col);
            ClassicAssert.IsTrue(ref1.IsRowAbsolute, "row absolute");
            //ClassicAssert.IsFalse("column absolute/relative is undefined", ref.IsColAbsolute);
        }

        [Test]
        public void GetSheetName()
        {
            ClassicAssert.AreEqual(null, new CellReference("A5").SheetName);
            ClassicAssert.AreEqual(null, new CellReference(null, 0, 0, false, false).SheetName);
            // FIXME: CellReference is inconsistent
            ClassicAssert.AreEqual("", new CellReference("", 0, 0, false, false).SheetName);
            ClassicAssert.AreEqual("Sheet1", new CellReference("Sheet1!A5").SheetName);
            ClassicAssert.AreEqual("Sheet 1", new CellReference("'Sheet 1'!A5").SheetName);
        }

        [Test]
        public void TestToString()
        {
            CellReference ref1 = new CellReference("'Sheet 1'!A5");
            ClassicAssert.AreEqual(ref1.ToString(), "CellReference ['Sheet 1'!A5]");
        }

        [Test]
        public void TestEqualsAndHashCode()
        {
            CellReference ref1 = new CellReference("'Sheet 1'!A5");
            CellReference ref2 = new CellReference("Sheet 1", 4, 0, false, false);
            ClassicAssert.AreEqual(ref1, ref2, "equals");
            ClassicAssert.AreEqual(ref1.GetHashCode(), ref2.GetHashCode(), "hash code");

            ClassicAssert.IsFalse(ref1.Equals(null), "null");
            ClassicAssert.IsFalse(ref1.Equals(new CellReference("A5")), "3D vs 2D");
            ClassicAssert.IsFalse(ref1.Equals(0), "type");
        }

        [Test]
        public void IsRowWithinRange()
        {
            SpreadsheetVersion ss = SpreadsheetVersion.EXCEL2007;
            ClassicAssert.IsFalse(CellReference.IsRowWithinRange("0", ss), "1 before first row");
            ClassicAssert.IsTrue(CellReference.IsRowWithinRange("1", ss), "first row");
            ClassicAssert.IsTrue(CellReference.IsRowWithinRange("1048576", ss), "last row");
            ClassicAssert.IsFalse(CellReference.IsRowWithinRange("1048577", ss), "1 beyond last row");

            ClassicAssert.IsFalse(CellReference.IsRowWithinRange(-1, ss), "1 before first row");
            ClassicAssert.IsTrue(CellReference.IsRowWithinRange(0, ss), "first row");
            ClassicAssert.IsTrue(CellReference.IsRowWithinRange(1048575, ss), "last row");
            ClassicAssert.IsFalse(CellReference.IsRowWithinRange(1048576, ss), "1 beyond last row");
        }

        [Test] //(expected= NumberFormatException.class)
        public void IsRowWithinRangeNonInteger_BigNumber()
        {
            String rowNum = "4000000000";
            CellReference.IsRowWithinRange(rowNum, SpreadsheetVersion.EXCEL2007);
        }

        [Test] //(expected= NumberFormatException.class)
        public void IsRowWithinRangeNonInteger_Alpha()
        {
            String rowNum = "NotANumber";
            CellReference.IsRowWithinRange(rowNum, SpreadsheetVersion.EXCEL2007);
         }

        [Test]
        public void IsColWithinRange()
        {
            SpreadsheetVersion ss = SpreadsheetVersion.EXCEL2007;
            ClassicAssert.IsTrue(CellReference.IsColumnWithinRange("", ss), "(empty)");
            ClassicAssert.IsTrue(CellReference.IsColumnWithinRange("A", ss), "first column (A)");
            ClassicAssert.IsTrue(CellReference.IsColumnWithinRange("XFD", ss), "last column (XFD)");
            ClassicAssert.IsFalse(CellReference.IsColumnWithinRange("XFE", ss), "1 beyond last column (XFE)");
        }

        [Test]
        public void UnquotedSheetName()
        {
            Assert.Throws(typeof(ArgumentException), () => {
                new CellReference("'Sheet 1!A5");
            });
        }
        [Test]
        public void MismatchedQuotesSheetName()
        {
            Assert.Throws(typeof(ArgumentException), () => {
                new CellReference("Sheet 1!A5");
            });
        }

        [Test]
        public void EscapedSheetName()
        {
            String escapedName = "'Don''t Touch'!A5";
            String unescapedName = "'Don't Touch'!A5";
            new CellReference(escapedName);
            try
            {
                new CellReference(unescapedName);
                Assert.Fail("Sheet names containing apostrophe's must be escaped via a repeated apostrophe");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.IsTrue(e.Message.StartsWith("Bad sheet name quote escaping: "));
            }
        }

        [Test]
        public void NegativeRow()
        {
            Assert.Throws(typeof(ArgumentException), () => {
                new CellReference("sheet", -2, 0, false, false);
            });
        }
        [Test]
        public void NegativeColumn()
        {
            Assert.Throws(typeof(ArgumentException), () => {
                new CellReference("sheet", 0, -2, false, false);
            });
        }

        [Test]
        public void ClassifyEmptyStringCellReference()
        {
            Assert.Throws(typeof(ArgumentException), () => {
                CellReference.ClassifyCellReference("", SpreadsheetVersion.EXCEL2007);
            });
        }
        [Test]
        public void ClassifyInvalidFirstCharCellReference()
        {
            Assert.Throws(typeof(ArgumentException), () => {
                CellReference.ClassifyCellReference("!A5", SpreadsheetVersion.EXCEL2007);
            }); 
        }

    }

}