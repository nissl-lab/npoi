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

    using NPOI.SS;
    using TestCases.SS;
    using NPOI.SS.UserModel;

    /**
     */
    [TestFixture]
    public class BaseTestSheetHiding
    {

        protected ITestDataProvider _testDataProvider;
        protected IWorkbook wbH;
        protected IWorkbook wbU;

        private String _file1, _file2;
        public BaseTestSheetHiding()
            : this(TestCases.HSSF.HSSFITestDataProvider.Instance, "TwoSheetsOneHidden.xls", "TwoSheetsNoneHidden.xls")
        {
        }
        /**
         * @param TestDataProvider an object that provides Test data in HSSF /  specific way
         */
        protected BaseTestSheetHiding(ITestDataProvider TestDataProvider,
                                      String file1, String file2)
        {
            _testDataProvider = TestDataProvider;
            _file1 = file1;
            _file2 = file2;
        }
        [SetUp]
        public void SetUp()
        {
            wbH = _testDataProvider.OpenSampleWorkbook(_file1);
            wbU = _testDataProvider.OpenSampleWorkbook(_file2);
        }

        [Test]
        [Obsolete]
        public void TestSheetHiddenOld()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("MySheet");

            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));

            wb.SetSheetHidden(0, SheetVisibility.Hidden);
            Assert.IsTrue(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));

            wb.SetSheetHidden(0, SheetVisibility.VeryHidden);
            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsTrue(wb.IsSheetVeryHidden(0));

            wb.SetSheetHidden(0, SheetVisibility.Visible);
            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));

            try
            {
                wb.SetSheetHidden(0, -1);
                Assert.Fail("expectd exception");
            }
            catch (ArgumentException)
            {
                // ok
            }
            try
            {
                wb.SetSheetHidden(0, 3);
                Assert.Fail("expectd exception");
            }
            catch (ArgumentException)
            {
                // ok
            }

            wb.Close();
        }

        [Test]
        public void TestSheetVisibility()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("MySheet");

            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));
            Assert.AreEqual(SheetVisibility.Visible, wb.GetSheetVisibility(0));

            wb.SetSheetVisibility(0, SheetVisibility.Hidden);
            Assert.IsTrue(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));
            Assert.AreEqual(SheetVisibility.Hidden, wb.GetSheetVisibility(0));

            wb.SetSheetVisibility(0, SheetVisibility.VeryHidden);
            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsTrue(wb.IsSheetVeryHidden(0));
            Assert.AreEqual(SheetVisibility.VeryHidden, wb.GetSheetVisibility(0));

            wb.SetSheetVisibility(0, SheetVisibility.Visible);
            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));
            Assert.AreEqual(SheetVisibility.Visible, wb.GetSheetVisibility(0));

            wb.Close();
        }

        /**
         * Test that we Get the right number of sheets,
         *  with the right text on them, no matter what
         *  the hidden flags are
         */
        [Test]
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
         * Check that we can Get and Set the hidden flags
         *  as expected
         */
        [Test]
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
        [Test]
        public void TestHide()
        {
            wbU.SetSheetHidden(0,SheetVisibility.Hidden);
            Assert.IsTrue(wbU.IsSheetHidden(0));
            Assert.IsFalse(wbU.IsSheetHidden(1));
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wbU);
            Assert.IsTrue(wb2.IsSheetHidden(0));
            Assert.IsFalse(wb2.IsSheetHidden(1));

            wb2.Close();
        }

        /**
         * Turn the sheet with one hidden into the one with
         *  none hidden
         */
        [Test]
        public void TestUnHide()
        {
            wbH.SetSheetHidden(0,SheetVisibility.Visible);
            Assert.IsFalse(wbH.IsSheetHidden(0));
            Assert.IsFalse(wbH.IsSheetHidden(1));
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wbH);
            Assert.IsFalse(wb2.IsSheetHidden(0));
            Assert.IsFalse(wb2.IsSheetHidden(1));

            wb2.Close();
        }
    }
}