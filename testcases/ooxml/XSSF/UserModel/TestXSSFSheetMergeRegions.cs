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
namespace NPOI.XSSF.UserModel
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using NPOI.SS.Util;
    using NPOI.XSSF;
    using NUnit.Framework;

    [TestFixture]
    public class TestXSSFSheetMergeRegions
    {
        [Test]
        public void TestMergeRegionsSpeed()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57893-many-merges.xlsx");
            try
            {
                XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
                Stopwatch stopwatch = Stopwatch.StartNew();
                //long start = System.CurrentTimeMillis();
                List<CellRangeAddress> mergedRegions = sheet.MergedRegions;
                Assert.AreEqual(50000, mergedRegions.Count);
                foreach (CellRangeAddress cellRangeAddress in mergedRegions)
                {
                    Assert.AreEqual(cellRangeAddress.FirstRow, cellRangeAddress.LastRow);
                    Assert.AreEqual(2, cellRangeAddress.NumberOfCells);
                }
                //long millis = System.CurrentTimeMillis() - start;
                stopwatch.Stop();
                long millis = stopwatch.ElapsedMilliseconds;
                // This time is typically ~800ms, versus ~7800ms to iterate GetMergedRegion(int).
                Assert.IsTrue(millis < 2000, "Should have taken <2000 ms to iterate 50k merged regions but took " + millis);
            }
            finally
            {
                wb.Close();
            }
        }
    }
}
