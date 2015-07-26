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

namespace TestCases.HSSF.Util
{
    using System;


    using NUnit.Framework;

    using NPOI.SS.Util;
    using NPOI.SS;


    [TestFixture]
    public class TestCellReference
    {
        [Test]
        public void TestColNumConversion()
        {
            Assert.AreEqual(0, CellReference.ConvertColStringToIndex("A"));
            Assert.AreEqual(1, CellReference.ConvertColStringToIndex("B"));
            Assert.AreEqual(25, CellReference.ConvertColStringToIndex("Z"));
            Assert.AreEqual(26, CellReference.ConvertColStringToIndex("AA"));
            Assert.AreEqual(27, CellReference.ConvertColStringToIndex("AB"));
            Assert.AreEqual(51, CellReference.ConvertColStringToIndex("AZ"));
            Assert.AreEqual(701, CellReference.ConvertColStringToIndex("ZZ"));
            Assert.AreEqual(702, CellReference.ConvertColStringToIndex("AAA"));
            Assert.AreEqual(18277, CellReference.ConvertColStringToIndex("ZZZ"));

            Assert.AreEqual("A", CellReference.ConvertNumToColString(0));
            Assert.AreEqual("B", CellReference.ConvertNumToColString(1));
            Assert.AreEqual("Z", CellReference.ConvertNumToColString(25));
            Assert.AreEqual("AA", CellReference.ConvertNumToColString(26));
            Assert.AreEqual("ZZ", CellReference.ConvertNumToColString(701));
            Assert.AreEqual("AAA", CellReference.ConvertNumToColString(702));
            Assert.AreEqual("ZZZ", CellReference.ConvertNumToColString(18277));

            // Absolute references are allowed for the string ones
            Assert.AreEqual(0, CellReference.ConvertColStringToIndex("$A"));
            Assert.AreEqual(25, CellReference.ConvertColStringToIndex("$Z"));
            Assert.AreEqual(26, CellReference.ConvertColStringToIndex("$AA"));

            // $ sign isn't allowed elsewhere though
            try
            {
                CellReference.ConvertColStringToIndex("A$B$");
                Assert.Fail("Column reference is invalid and shouldn't be accepted");
            }
            catch (ArgumentException) { }
        }

        [Test]
        public void TestAbsRef1()
        {
            CellReference cf = new CellReference("$B$5");
            ConfirmCell(cf, null, 4, 1, true, true, "$B$5");
        }
        [Test]
        public void TestAbsRef2()
        {
            CellReference cf = new CellReference(4, 1, true, true);
            ConfirmCell(cf, null, 4, 1, true, true, "$B$5");
        }
        [Test]
        public void TestAbsRef3()
        {
            CellReference cf = new CellReference("B$5");
            ConfirmCell(cf, null, 4, 1, true, false, "B$5");
        }
        [Test]
        public void TestAbsRef4()
        {
            CellReference cf = new CellReference(4, 1, true, false);
            ConfirmCell(cf, null, 4, 1, true, false, "B$5");
        }
        [Test]
        public void TestAbsRef5()
        {
            CellReference cf = new CellReference("$B5");
            ConfirmCell(cf, null, 4, 1, false, true, "$B5");
        }
        [Test]
        public void TestAbsRef6()
        {
            CellReference cf = new CellReference(4, 1, false, true);
            ConfirmCell(cf, null, 4, 1, false, true, "$B5");
        }
        [Test]
        public void TestAbsRef7()
        {
            CellReference cf = new CellReference("B5");
            ConfirmCell(cf, null, 4, 1, false, false, "B5");
        }
        [Test]
        public void TestAbsRef8()
        {
            CellReference cf = new CellReference(4, 1, false, false);
            ConfirmCell(cf, null, 4, 1, false, false, "B5");
        }
        [Test]
        public void TestSpecialSheetNames()
        {
            CellReference cf;
            cf = new CellReference("'profit + loss'!A1");
            ConfirmCell(cf, "profit + loss", 0, 0, false, false, "'profit + loss'!A1");

            cf = new CellReference("'O''Brien''s Sales'!A1");
            ConfirmCell(cf, "O'Brien's Sales", 0, 0, false, false, "'O''Brien''s Sales'!A1");

            cf = new CellReference("'Amazing!'!A1");
            ConfirmCell(cf, "Amazing!", 0, 0, false, false, "'Amazing!'!A1");
        }

        /* package */
        internal static void ConfirmCell(CellReference cf, String expSheetName, int expRow,
            int expCol, bool expIsRowAbs, bool expIsColAbs, String expText)
        {

            Assert.AreEqual(expSheetName, cf.SheetName);
            Assert.AreEqual(expRow, cf.Row, "row index is wrong");
            Assert.AreEqual(expCol, cf.Col, "col index is wrong");
            Assert.AreEqual(expIsRowAbs, cf.IsRowAbsolute, "isRowAbsolute is wrong");
            Assert.AreEqual(expIsColAbs, cf.IsColAbsolute, "isColAbsolute is wrong");
            Assert.AreEqual(expText, cf.FormatAsString(), "text is wrong");
        }
        [Test]
        public void TestClassifyCellReference()
        {
            ConfirmNameType("a1", NameType.Cell);
            ConfirmNameType("pfy1", NameType.NamedRange);
            ConfirmNameType("pf1", NameType.NamedRange); // (col) out of cell range
            ConfirmNameType("fp1", NameType.Cell);
            ConfirmNameType("pf$1", NameType.BadCellOrNamedRange);
            ConfirmNameType("_A1", NameType.NamedRange);
            ConfirmNameType("A_1", NameType.NamedRange);
            ConfirmNameType("A1_", NameType.NamedRange);
            ConfirmNameType(".A1", NameType.BadCellOrNamedRange);
            ConfirmNameType("A.1", NameType.NamedRange);
            ConfirmNameType("A1.", NameType.NamedRange);
        }
        [Test]
        public void TestClassificationOfRowReferences()
        {
            ConfirmNameType("10", NameType.Row);
            ConfirmNameType("$10", NameType.Row);
            ConfirmNameType("65536", NameType.Row);

            ConfirmNameType("65537", NameType.BadCellOrNamedRange);
            ConfirmNameType("$100000", NameType.BadCellOrNamedRange);
            ConfirmNameType("$1$1", NameType.BadCellOrNamedRange);
        }

        private void ConfirmNameType(String ref1, NameType expectedResult)
        {
            NameType actualResult = CellReference.ClassifyCellReference(ref1, SpreadsheetVersion.EXCEL97);
            Assert.AreEqual(expectedResult, actualResult);
        }

        [Test]
        public void TestConvertColStringToIndex()
        {
            Assert.AreEqual(0, CellReference.ConvertColStringToIndex("A"));
            Assert.AreEqual(1, CellReference.ConvertColStringToIndex("B"));
            Assert.AreEqual(14, CellReference.ConvertColStringToIndex("O"));
            Assert.AreEqual(701, CellReference.ConvertColStringToIndex("ZZ"));
            Assert.AreEqual(18252, CellReference.ConvertColStringToIndex("ZZA"));

            Assert.AreEqual(0, CellReference.ConvertColStringToIndex("$A"));
            Assert.AreEqual(1, CellReference.ConvertColStringToIndex("$B"));

            try
            {
                CellReference.ConvertColStringToIndex("A$");
                Assert.Fail("Should throw exception here");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.Contains("A$"));
            }
        }

        [Test]
        public void TestConvertNumColColString()
        {
            Assert.AreEqual("A", CellReference.ConvertNumToColString(0));
            Assert.AreEqual("AV", CellReference.ConvertNumToColString(47));
            Assert.AreEqual("AW", CellReference.ConvertNumToColString(48));
            Assert.AreEqual("BF", CellReference.ConvertNumToColString(57));

            Assert.AreEqual("", CellReference.ConvertNumToColString(-1));
            Assert.AreEqual("", CellReference.ConvertNumToColString(Int32.MinValue));
            Assert.AreEqual("", CellReference.ConvertNumToColString(Int32.MaxValue));
            Assert.AreEqual("FXSHRXW", CellReference.ConvertNumToColString(Int32.MaxValue - 1));
        }

    }

}