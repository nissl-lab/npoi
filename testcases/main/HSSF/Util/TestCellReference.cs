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
    }

}