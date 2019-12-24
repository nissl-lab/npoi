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

namespace TestCases.SS.Util
{
    using System;

    using NPOI.SS.Util;
    using NPOI.Util;
    using NUnit.Framework;




    /**
     * Tests that the common CellAddress works as we need it to.
     * Note - some Additional testing is also done in the HSSF class,
     *  {@link NPOI.HSSF.Util.TestCellAddress}
     */
    [TestFixture]
    public class TestCellAddress
    {
        [Test]
        public void TestConstructors()
        {
            CellAddress cellAddress;
            CellReference cellRef = new CellReference("Sheet1", 0, 0, true, true);
            String Address = "A1";
            int row = 0;
            int col = 0;

            cellAddress = new CellAddress(row, col);
            Assert.AreEqual(CellAddress.A1, cellAddress);

            cellAddress = new CellAddress(Address);
            Assert.AreEqual(CellAddress.A1, cellAddress);

            cellAddress = new CellAddress(cellRef);
            Assert.AreEqual(CellAddress.A1, cellAddress);
        }

        [Test]
        public void TestFormatAsString()
        {
            Assert.AreEqual("A1", CellAddress.A1.FormatAsString());
        }

        [Test]
        public void TestEquals()
        {
            Assert.AreEqual(new CellReference(6, 4), new CellReference(6, 4));
            Assert.AreNotEqual(new CellReference(4, 6), new CellReference(6, 4));
        }

        [Test]
        public void TestCompareTo()
        {
            CellAddress A1 = new CellAddress(0, 0);
            CellAddress A2 = new CellAddress(1, 0);
            CellAddress B1 = new CellAddress(0, 1);
            CellAddress B2 = new CellAddress(1, 1);

            Assert.AreEqual(0, A1.CompareTo(A1));
            Assert.AreEqual(-1, A1.CompareTo(B1));
            Assert.AreEqual(-1, A1.CompareTo(A2));
            Assert.AreEqual(-1, A1.CompareTo(B2));

            Assert.AreEqual(1, B1.CompareTo(A1));
            Assert.AreEqual(0, B1.CompareTo(B1));
            Assert.AreEqual(-1, B1.CompareTo(A2));
            Assert.AreEqual(-1, B1.CompareTo(B2));

            Assert.AreEqual(1, A2.CompareTo(A1));
            Assert.AreEqual(1, A2.CompareTo(B1));
            Assert.AreEqual(0, A2.CompareTo(A2));
            Assert.AreEqual(-1, A2.CompareTo(B2));

            Assert.AreEqual(1, B2.CompareTo(A1));
            Assert.AreEqual(1, B2.CompareTo(B1));
            Assert.AreEqual(1, B2.CompareTo(A2));
            Assert.AreEqual(0, B2.CompareTo(B2));

            CellAddress[] sorted = { A1, B1, A2, B2 };
            CellAddress[] unsorted = { B1, B2, A1, A2 };
            Assume.That(!sorted.Equals(unsorted));
            Array.Sort(unsorted);
            CollectionAssert.AreEqual(sorted, unsorted);
        }
        [Test]
        public void TestGetRow()
        {
            CellAddress Addr = new CellAddress(6, 4);
            Assert.AreEqual(6, Addr.Row);
        }

        [Test]
        public void TestGetColumn()
        {
            CellAddress Addr = new CellAddress(6, 4);
            Assert.AreEqual(4, Addr.Column);
        }

    }
}