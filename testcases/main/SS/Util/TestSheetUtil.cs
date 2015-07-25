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

    using NUnit.Framework;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * Tests SheetUtil.
     *
     * @see NPOI.SS.Util.SheetUtil
     */
    [TestFixture]
    public class TestSheetUtil
    {
        [Test]
        public void TestCellWithMerges()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet();

            // Create some test data
            IRow r2 = s.CreateRow(1);
            r2.CreateCell(0).SetCellValue(10);
            r2.CreateCell(1).SetCellValue(11);
            IRow r3 = s.CreateRow(2);
            r3.CreateCell(0).SetCellValue(20);
            r3.CreateCell(1).SetCellValue(21);

            s.AddMergedRegion(new CellRangeAddress(2, 3, 0, 0));
            s.AddMergedRegion(new CellRangeAddress(2, 2, 1, 4));

            // With a cell that isn't defined, we'll Get null
            Assert.AreEqual(null, SheetUtil.GetCellWithMerges(s, 0, 0));

            // With a cell that's not in a merged region, we'll Get that
            Assert.AreEqual(10.0, SheetUtil.GetCellWithMerges(s, 1, 0).NumericCellValue);
            Assert.AreEqual(11.0, SheetUtil.GetCellWithMerges(s, 1, 1).NumericCellValue);

            // With a cell that's the primary one of a merged region, we Get that cell
            Assert.AreEqual(20.0, SheetUtil.GetCellWithMerges(s, 2, 0).NumericCellValue);
            Assert.AreEqual(21, SheetUtil.GetCellWithMerges(s, 2, 1).NumericCellValue);

            // With a cell elsewhere in the merged region, Get top-left
            Assert.AreEqual(20.0, SheetUtil.GetCellWithMerges(s, 3, 0).NumericCellValue);
            Assert.AreEqual(21.0, SheetUtil.GetCellWithMerges(s, 2, 2).NumericCellValue);
            Assert.AreEqual(21.0, SheetUtil.GetCellWithMerges(s, 2, 3).NumericCellValue);
            Assert.AreEqual(21.0, SheetUtil.GetCellWithMerges(s, 2, 4).NumericCellValue);
        }
    }
}