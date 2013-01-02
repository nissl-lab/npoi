
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


namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    /**
     * Tests HSSFWorkbook method setSheetOrder()
     *
     *
     * @author Ruel Loehr (loehr1 at us.ibm.com)
     */
    [TestFixture]
    public class TestHSSFSheetOrder
    {
        public TestHSSFSheetOrder()
        {

        }

        /**
         * Test the sheet set order method
         */
        [Test]
        public void TestBackupRecord()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            for (int i = 0; i < 10; i++)
            {
                HSSFSheet s = (HSSFSheet)wb.CreateSheet("Sheet " + i);
                InternalSheet sheet = s.Sheet;
            }

            // Check the initial order
            Assert.AreEqual(0, wb.GetSheetIndex("Sheet 0"));
            Assert.AreEqual(1, wb.GetSheetIndex("Sheet 1"));
            Assert.AreEqual(2, wb.GetSheetIndex("Sheet 2"));
            Assert.AreEqual(3, wb.GetSheetIndex("Sheet 3"));
            Assert.AreEqual(4, wb.GetSheetIndex("Sheet 4"));
            Assert.AreEqual(5, wb.GetSheetIndex("Sheet 5"));
            Assert.AreEqual(6, wb.GetSheetIndex("Sheet 6"));
            Assert.AreEqual(7, wb.GetSheetIndex("Sheet 7"));
            Assert.AreEqual(8, wb.GetSheetIndex("Sheet 8"));
            Assert.AreEqual(9, wb.GetSheetIndex("Sheet 9"));

            // Change
            wb.Workbook.SetSheetOrder("Sheet 6", 0);
            wb.Workbook.SetSheetOrder("Sheet 3", 7);
            wb.Workbook.SetSheetOrder("Sheet 1", 9);

            // Check they're currently right
            Assert.AreEqual(0, wb.GetSheetIndex("Sheet 6"));
            Assert.AreEqual(1, wb.GetSheetIndex("Sheet 0"));
            Assert.AreEqual(2, wb.GetSheetIndex("Sheet 2"));
            Assert.AreEqual(3, wb.GetSheetIndex("Sheet 4"));
            Assert.AreEqual(4, wb.GetSheetIndex("Sheet 5"));
            Assert.AreEqual(5, wb.GetSheetIndex("Sheet 7"));
            Assert.AreEqual(6, wb.GetSheetIndex("Sheet 3"));
            Assert.AreEqual(7, wb.GetSheetIndex("Sheet 8"));
            Assert.AreEqual(8, wb.GetSheetIndex("Sheet 9"));
            Assert.AreEqual(9, wb.GetSheetIndex("Sheet 1"));

            // Read it in and see if it is correct.
            MemoryStream baos = new MemoryStream();
            wb.Write(baos);
            MemoryStream bais = new MemoryStream(baos.ToArray());
            HSSFWorkbook wbr = new HSSFWorkbook(bais);

            Assert.AreEqual(0, wbr.GetSheetIndex("Sheet 6"));
            Assert.AreEqual(1, wbr.GetSheetIndex("Sheet 0"));
            Assert.AreEqual(2, wbr.GetSheetIndex("Sheet 2"));
            Assert.AreEqual(3, wbr.GetSheetIndex("Sheet 4"));
            Assert.AreEqual(4, wbr.GetSheetIndex("Sheet 5"));
            Assert.AreEqual(5, wbr.GetSheetIndex("Sheet 7"));
            Assert.AreEqual(6, wbr.GetSheetIndex("Sheet 3"));
            Assert.AreEqual(7, wbr.GetSheetIndex("Sheet 8"));
            Assert.AreEqual(8, wbr.GetSheetIndex("Sheet 9"));
            Assert.AreEqual(9, wbr.GetSheetIndex("Sheet 1"));

            // Now get the index by the sheet, not the name
            for (int i = 0; i < 10; i++)
            {
                NPOI.SS.UserModel.ISheet s = wbr.GetSheetAt(i);
                Assert.AreEqual(i, wbr.GetSheetIndex(s));
            }
        }
    }
}

