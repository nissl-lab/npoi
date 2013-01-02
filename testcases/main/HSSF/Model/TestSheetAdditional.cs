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

namespace TestCases.HSSF.Model
{
    using System;
    using System.Collections;

    using NUnit.Framework;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Model;


    /**
     * @author Tony Poppleton
     */
    [TestFixture]
    public class TestSheetAdditional
    {
        [Test]
        public void TestGetCellWidth()
        {
            InternalSheet sheet = InternalSheet.CreateSheet();
            ColumnInfoRecord nci = new ColumnInfoRecord();

            // Prepare test model
            nci.FirstColumn = 5;
            nci.LastColumn = 10;
            nci.ColumnWidth = 100;


            sheet._columnInfos.InsertColumn(nci);

            Assert.AreEqual(100, sheet.GetColumnWidth(5));
            Assert.AreEqual(100, sheet.GetColumnWidth(6));
            Assert.AreEqual(100, sheet.GetColumnWidth(7));
            Assert.AreEqual(100, sheet.GetColumnWidth(8));
            Assert.AreEqual(100, sheet.GetColumnWidth(9));
            Assert.AreEqual(100, sheet.GetColumnWidth(10));

            sheet.SetColumnWidth(6, 200);

            Assert.AreEqual(100, sheet.GetColumnWidth(5));
            Assert.AreEqual(200, sheet.GetColumnWidth(6));
            Assert.AreEqual(100, sheet.GetColumnWidth(7));
            Assert.AreEqual(100, sheet.GetColumnWidth(8));
            Assert.AreEqual(100, sheet.GetColumnWidth(9));
            Assert.AreEqual(100, sheet.GetColumnWidth(10));
        }
        [Test]
        public void TestMaxColumnWidth()
        {
            InternalSheet sheet = InternalSheet.CreateSheet();
            sheet.SetColumnWidth(0, 255 * 256); //the limit
            try
            {
                sheet.SetColumnWidth(0, 256 * 256); //the limit
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual(e.Message, "The maximum column width for an individual cell is 255 characters.");
            }
        }
    }
}