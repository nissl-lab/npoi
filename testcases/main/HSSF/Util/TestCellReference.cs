/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Util
{

    using System;
    using NPOI.HSSF.Util;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.Util;

    [TestClass]
    public class TestCellReference
    {
        [TestMethod]
        public void TestAbsRef1()
        {
            CellReference cf = new CellReference("$B$5");
            ConfirmCell(cf, null, 4, 1, true, true, "$B$5");
        }
        [TestMethod]
        public void TestAbsRef2()
        {
            CellReference cf = new CellReference(4, 1, true, true);
            ConfirmCell(cf, null, 4, 1, true, true, "$B$5");
        }
        [TestMethod]
        public void TestAbsRef3()
        {
            CellReference cf = new CellReference("B$5");
            ConfirmCell(cf, null, 4, 1, true, false, "B$5");
        }
        [TestMethod]
        public void TestAbsRef4()
        {
            CellReference cf = new CellReference(4, 1, true, false);
            ConfirmCell(cf, null, 4, 1, true, false, "B$5");
        }
        [TestMethod]
        public void TestAbsRef5()
        {
            CellReference cf = new CellReference("$B5");
            ConfirmCell(cf, null, 4, 1, false, true, "$B5");
        }
        [TestMethod]
        public void TestAbsRef6()
        {
            CellReference cf = new CellReference(4, 1, false, true);
            ConfirmCell(cf, null, 4, 1, false, true, "$B5");
        }
        [TestMethod]
        public void TestAbsRef7()
        {
            CellReference cf = new CellReference("B5");
            ConfirmCell(cf, null, 4, 1, false, false, "B5");
        }
        [TestMethod]
        public void TestAbsRef8()
        {
            CellReference cf = new CellReference(4, 1, false, false);
            ConfirmCell(cf, null, 4, 1, false, false, "B5");
        }
        [TestMethod]
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

        public static void ConfirmCell(CellReference cf, String expSheetName, int expRow,
                int expCol, bool expIsRowAbs, bool expIsColAbs, String expText)
        {

            Assert.AreEqual(expSheetName, cf.SheetName);
            Assert.AreEqual(expRow, cf.Row, "row index is1 wrong");
            Assert.AreEqual(expCol, cf.Col, "col index is1 wrong");
            Assert.AreEqual(expIsRowAbs, cf.IsRowAbsolute, "isRowAbsolute is1 wrong");
            Assert.AreEqual(expIsColAbs, cf.IsColAbsolute, "isColAbsolute is1 wrong");
            Assert.AreEqual(expText, cf.FormatAsString(), "text is1 wrong");
        }
        [TestMethod]
        public void TestClassifyCellReference()
        {
            ConfirmNameType("a1", NameType.CELL);
            ConfirmNameType("pfy1", NameType.NAMED_RANGE);
            ConfirmNameType("pf1", NameType.NAMED_RANGE); // (col) out of cell range
            ConfirmNameType("fp1", NameType.CELL);
            ConfirmNameType("pf$1", NameType.BAD_CELL_OR_NAMED_RANGE);
            ConfirmNameType("_A1", NameType.NAMED_RANGE);
            ConfirmNameType("A_1", NameType.NAMED_RANGE);
            ConfirmNameType("A1_", NameType.NAMED_RANGE);
            ConfirmNameType(".A1", NameType.BAD_CELL_OR_NAMED_RANGE);
            ConfirmNameType("A.1", NameType.NAMED_RANGE);
            ConfirmNameType("A1.", NameType.NAMED_RANGE);
        }

        private void ConfirmNameType(String ref1, NameType expectedResult)
        {
            NameType actualResult = CellReference.ClassifyCellReference(ref1, NPOI.SS.SpreadsheetVersion.EXCEL97);
            Assert.AreEqual(expectedResult, actualResult);
        }

        [TestMethod]
        public void TestCellRefParts()
        {
            CellReference cellReference;
            String[] parts;

            String cellRef = "A1";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(0, cellReference.Col);
            parts = cellReference.CellRefParts;
            Assert.IsNotNull(parts);
            Assert.AreEqual(null, parts[0]);
            Assert.AreEqual("1", parts[1]);
            Assert.AreEqual("A", parts[2]);

            cellRef = "AA1";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26, cellReference.Col);
            parts = cellReference.CellRefParts;
            Assert.IsNotNull(parts);
            Assert.AreEqual(null, parts[0]);
            Assert.AreEqual("1", parts[1]);
            Assert.AreEqual("AA", parts[2]);

            cellRef = "AA100";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26, cellReference.Col);
            parts = cellReference.CellRefParts;
            Assert.IsNotNull(parts);
            Assert.AreEqual(null, parts[0]);
            Assert.AreEqual("100", parts[1]);
            Assert.AreEqual("AA", parts[2]);

            cellRef = "AAA300";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(702, cellReference.Col);
            parts = cellReference.CellRefParts;
            Assert.IsNotNull(parts);
            Assert.AreEqual(null, parts[0]);
            Assert.AreEqual("300", parts[1]);
            Assert.AreEqual("AAA", parts[2]);

            cellRef = "ZZ100521";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26 * 26 + 25, cellReference.Col);
            parts = cellReference.CellRefParts;
            Assert.IsNotNull(parts);
            Assert.AreEqual(null, parts[0]);
            Assert.AreEqual("100521", parts[1]);
            Assert.AreEqual("ZZ", parts[2]);

            cellRef = "ZYX987";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26 * 26 * 26 + 25 * 26 + 24 - 1, cellReference.Col);
            parts = cellReference.CellRefParts;
            Assert.IsNotNull(parts);
            Assert.AreEqual(null, parts[0]);
            Assert.AreEqual("987", parts[1]);
            Assert.AreEqual("ZYX", parts[2]);

            cellRef = "AABC10065";
            cellReference = new CellReference(cellRef);
            parts = cellReference.CellRefParts;
            Assert.IsNotNull(parts);
            Assert.AreEqual(null, parts[0]);
            Assert.AreEqual("10065", parts[1]);
            Assert.AreEqual("AABC", parts[2]);
        }
        [TestMethod]
        public void TestGetColNumFromRef()
        {
            String cellRef = "A1";
            CellReference cellReference = new CellReference(cellRef);
            Assert.AreEqual(0, cellReference.Col);

            cellRef = "AA1";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26, cellReference.Col);

            cellRef = "AB1";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(27, cellReference.Col);

            cellRef = "BA1";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26 + 26, cellReference.Col);

            cellRef = "CA1";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26 + 26 + 26, cellReference.Col);

            cellRef = "ZA1";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26 * 26, cellReference.Col);

            cellRef = "ZZ1";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26 * 26 + 25, cellReference.Col);

            cellRef = "AAA1";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(26 * 26 + 26, cellReference.Col);


            cellRef = "A1100";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(0, cellReference.Col);

            cellRef = "BC15";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(54, cellReference.Col);
        }
        [TestMethod]
        public void TestGetRowNumFromRef()
        {
            String cellRef = "A1";
            CellReference cellReference = new CellReference(cellRef);
            Assert.AreEqual(0, cellReference.Row);

            cellRef = "A12";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(11, cellReference.Row);

            cellRef = "AS121";
            cellReference = new CellReference(cellRef);
            Assert.AreEqual(120, cellReference.Row);
        }
        [TestMethod]
        public void TestConvertNumToColString()
        {
            short col = 702;
            String collRef = new CellReference(0, col).FormatAsString();
            Assert.AreEqual("AAA1", collRef);

            short col2 = 0;
            String collRef2 = new CellReference(0, col2).FormatAsString();
            Assert.AreEqual("A1", collRef2);

            short col3 = 27;
            String collRef3 = new CellReference(0, col3).FormatAsString();
            Assert.AreEqual("AB1", collRef3);

            short col4 = 2080;
            String collRef4 = new CellReference(0, col4).FormatAsString();
            Assert.AreEqual("CBA1", collRef4);
        }
    }
}