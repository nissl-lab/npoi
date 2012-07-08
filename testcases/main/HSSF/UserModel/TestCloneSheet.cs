
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
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Util;
    using NUnit.Framework;
    using NPOI.SS.UserModel;
    /**
     * Test the ability to clone a sheet. 
     *  If Adding new records that belong to a sheet (as opposed to a book)
     *  Add that record to the sheet in the TestCloneSheetBasic method. 
     * @author  avik
     */
    [TestFixture]
    public class TestCloneSheet
    {
        [Test]
        public void TestCloneSheetBasic()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            ISheet s = b.CreateSheet("Test");
            s.AddMergedRegion(new CellRangeAddress(0, 1, 0, 1));
            ISheet clonedSheet = b.CloneSheet(0);
            Assert.AreEqual(1, clonedSheet.NumMergedRegions, "One merged area");
        }

        /**
         * Ensures that pagebreak cloning works properly
         *
         */
        [Test]
        public void TestPageBreakClones()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            ISheet s = b.CreateSheet("Test");
            s.SetRowBreak(3);
            s.SetColumnBreak((short)6);

            ISheet clone = b.CloneSheet(0);
            Assert.IsTrue(clone.IsRowBroken(3), "Row 3 not broken");
            Assert.IsTrue(clone.IsColumnBroken((short)6), "Column 6 not broken");

            s.RemoveRowBreak(3);

            Assert.IsTrue(clone.IsRowBroken(3), "Row 3 still should be broken");
        }
    }
}