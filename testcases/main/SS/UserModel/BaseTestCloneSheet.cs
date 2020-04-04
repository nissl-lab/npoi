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
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NUnit.Framework;
    using System;
    using TestCases.SS;

    /**
     * Common superclass for testing implementations of
     * Workbook.CloneSheet()
     */
    public abstract class BaseTestCloneSheet
    {

        private ITestDataProvider _testDataProvider;

        protected BaseTestCloneSheet(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }

        [Test]
        public void TestCloneSheetBasic()
        {
            IWorkbook b = _testDataProvider.CreateWorkbook();
            ISheet s = b.CreateSheet("Test");
            s.AddMergedRegion(new CellRangeAddress(0, 1, 0, 1));
            ISheet ClonedSheet = b.CloneSheet(0);

            Assert.AreEqual(1, ClonedSheet.NumMergedRegions, "One merged area");

            b.Close();
        }

        /**
         * Ensures that pagebreak cloning works properly
         * @throws IOException
         */
        [Test]
        public void TestPageBreakClones()
        {
            IWorkbook b = _testDataProvider.CreateWorkbook();
            ISheet s = b.CreateSheet("Test");
            s.SetRowBreak(3);
            s.SetColumnBreak((short)6);

            ISheet clone = b.CloneSheet(0);
            Assert.IsTrue(clone.IsRowBroken(3), "Row 3 not broken");
            Assert.IsTrue(clone.IsColumnBroken((short)6), "Column 6 not broken");

            s.RemoveRowBreak(3);

            Assert.IsTrue(clone.IsRowBroken(3), "Row 3 still should be broken");

            b.Close();
        }

        [Test]
        public void TestCloneSheetIntValid()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("Sheet01");
            wb.CloneSheet(0);
            Assert.AreEqual(2, wb.NumberOfSheets);
            try
            {
                wb.CloneSheet(2);
                Assert.Fail("ShouldFail");
            }
            catch (ArgumentException e)
            {
                // expected here
            }
        }

        [Test]
        public void TestCloneSheetIntInvalid()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet("Sheet01");
            try
            {
                wb.CloneSheet(1);
                Assert.Fail("Should Fail");
            }
            catch (ArgumentException e)
            {
                // expected here
            }
            Assert.AreEqual(1, wb.NumberOfSheets);
        }
    }

}