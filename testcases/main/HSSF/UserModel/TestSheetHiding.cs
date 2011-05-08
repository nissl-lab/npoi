/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/

namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using System.IO;
    using TestCases.HSSF;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /**
     * Tests for how HSSFWorkbook behaves with XLS files
     *  with a WORKBOOK directory entry (instead of the more
     *  usual, Workbook)
     */
    [TestClass]
    public class TestSheetHiding
    {
        private HSSFWorkbook wbH;
        private HSSFWorkbook wbU;

        [TestInitialize]
        public void SetUp()
        {
            wbH = HSSFTestDataSamples.OpenSampleWorkbook("TwoSheetsOneHidden.xls");
            wbU = HSSFTestDataSamples.OpenSampleWorkbook("TwoSheetsNoneHidden.xls");
        }

        /**
         * Test that we get the right number of sheets,
         *  with the right text on them, no matter what
         *  the hidden flags are
         */
        [TestMethod]
        public void TestTextSheets()
        {
            // Both should have two sheets
            Assert.AreEqual(2, wbH.NumberOfSheets);
            Assert.AreEqual(2, wbU.NumberOfSheets);

            // All sheets should have one row
            Assert.AreEqual(0, wbH.GetSheetAt(0).LastRowNum);
            Assert.AreEqual(0, wbH.GetSheetAt(1).LastRowNum);
            Assert.AreEqual(0, wbU.GetSheetAt(0).LastRowNum);
            Assert.AreEqual(0, wbU.GetSheetAt(1).LastRowNum);

            // All rows should have one column
            Assert.AreEqual(1, wbH.GetSheetAt(0).GetRow(0).LastCellNum);
            Assert.AreEqual(1, wbH.GetSheetAt(1).GetRow(0).LastCellNum);
            Assert.AreEqual(1, wbU.GetSheetAt(0).GetRow(0).LastCellNum);
            Assert.AreEqual(1, wbU.GetSheetAt(1).GetRow(0).LastCellNum);

            // Text should be sheet based
            Assert.AreEqual("Sheet1A1", wbH.GetSheetAt(0).GetRow(0).GetCell(0).RichStringCellValue.String);
            Assert.AreEqual("Sheet2A1", wbH.GetSheetAt(1).GetRow(0).GetCell(0).RichStringCellValue.String);
            Assert.AreEqual("Sheet1A1", wbU.GetSheetAt(0).GetRow(0).GetCell(0).RichStringCellValue.String);
            Assert.AreEqual("Sheet2A1", wbU.GetSheetAt(1).GetRow(0).GetCell(0).RichStringCellValue.String);
        }

        /**
         * Check that we can get and set the hidden flags
         *  as expected
         */
        [TestMethod]
        public void TestHideUnHideFlags()
        {
            Assert.IsTrue(wbH.IsSheetHidden(0));
            Assert.IsFalse(wbH.IsSheetHidden(1));
            Assert.IsFalse(wbU.IsSheetHidden(0));
            Assert.IsFalse(wbU.IsSheetHidden(1));
        }

        /**
         * Turn the sheet with none hidden into the one with
         *  one hidden
         */
        [TestMethod]
        public void TestHide()
        {
            wbU.SetSheetHidden(0, true);
            Assert.IsTrue(wbU.IsSheetHidden(0));
            Assert.IsFalse(wbU.IsSheetHidden(1));
            MemoryStream out1 = new MemoryStream();
            wbU.Write(out1);
            out1.Close();
            HSSFWorkbook wb2 = new HSSFWorkbook(new MemoryStream(out1.ToArray()));
            Assert.IsTrue(wb2.IsSheetHidden(0));
            Assert.IsFalse(wb2.IsSheetHidden(1));
        }

        /**
         * Turn the sheet with one hidden into the one with
         *  none hidden
         */
        [TestMethod]
        public void TestUnHide()
        {
            wbH.SetSheetHidden(0, false);
            Assert.IsFalse(wbH.IsSheetHidden(0));
            Assert.IsFalse(wbH.IsSheetHidden(1));
            MemoryStream out1 = new MemoryStream();
            wbH.Write(out1);
            out1.Close();
            HSSFWorkbook wb2 = new HSSFWorkbook(new MemoryStream(out1.ToArray()));
            Assert.IsFalse(wb2.IsSheetHidden(0));
            Assert.IsFalse(wb2.IsSheetHidden(1));
        }
    }
}