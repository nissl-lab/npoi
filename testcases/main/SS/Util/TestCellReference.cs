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
using NUnit.Framework;
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
        public void TestGetCellRefParts()
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
        [Test]
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
        [Test]
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
        [Test]
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
    }

}