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
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NUnit.Framework;

    [TestFixture]
    public class TestXSSFSheetRowGrouping
    {

        private static int ROWS_NUMBER = 200;
        private static int GROUP_SIZE = 5;

        //private int o_groupsNumber = 0;

        [Test]
        public void Test55640()
        {
            //long startTime = System.CurrentTimeMillis();
            IWorkbook wb = new XSSFWorkbook();
            FillData(wb);
            WriteToFile(wb);

            //System.out.Println("Number of groups: " + o_groupsNumber);
            //System.out.Println("Execution time: " + (System.CurrentTimeMillis()-startTime) + " ms");
        }


        private void FillData(IWorkbook p_wb)
        {
            ISheet sheet = p_wb.CreateSheet("sheet123");
            sheet.RowSumsBelow = (/*setter*/false);
            int i;
            for (i = 0; i < ROWS_NUMBER; i++)
            {
                IRow row = sheet.CreateRow(i);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(i + 1);
            }

            i = 1;
            while (i < ROWS_NUMBER)
            {
                int end = i + (GROUP_SIZE - 2);
                int start = i;                    // natural order
                //            int start = end - 1;                // reverse order
                while (start < end)
                {             // natural order
                    //                while (start >= i) {            // reverse order
                    sheet.GroupRow(start, end);
                    //o_groupsNumber++;
                    bool collapsed = IsCollapsed();
                    //System.out.Println("Set group " + start + "->" + end + " to " + collapsed);
                    sheet.SetRowGroupCollapsed(start, collapsed);
                    start++;                      // natural order
                    //                start--;                        // reverse order
                }
                i += GROUP_SIZE;
            }
        }

        private bool IsCollapsed()
        {
            Random rnd = new Random();
            return rnd.NextDouble() > 0.5d;
        }

        private void WriteToFile(IWorkbook p_wb)
        {
            //        FileOutputStream fileOut = new FileOutputStream("/tmp/55640.xlsx");
            //        try {
            //            p_wb.Write(fileOut);
            //        } finally {
            //            fileOut.Close();
            //        }
            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(p_wb));
        }

        [Test]
        public void Test55640reduce1()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet123");
            sheet.RowSumsBelow = (/*setter*/false);
            int i;
            for (i = 0; i < ROWS_NUMBER; i++)
            {
                IRow row = sheet.CreateRow(i);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(i + 1);
            }

            i = 1;
            while (i < ROWS_NUMBER)
            {
                int end = i + (GROUP_SIZE - 2);
                int start = i;                    // natural order
                while (start < end)
                {             // natural order
                    sheet.GroupRow(start, end);
                    //o_groupsNumber++;
                    bool collapsed = start % 2 == 0 ? false : true;
                    //System.out.Println("Set group " + start + "->" + end + " to " + collapsed);
                    sheet.SetRowGroupCollapsed(start, collapsed);
                    start++;                      // natural order
                }
                i += GROUP_SIZE;
            }
            WriteToFile(wb);
        }


        [Test]
        public void Test55640_VerifyCases()
        {
            // NOTE: This is currently based on current behavior of POI, somehow
            // what POI returns in the calls to collapsed/hidden is not fully matching 
            // the examples in the spec or I did not fully understand how POI stores the data internally...

            // all expanded
            verifyGroupCollapsed(
                // level1, level2, level3
                    false, false, false,
                // collapsed:
                    new Boolean[] { false, false, false, false, false },
                // hidden:
                    new bool[] { false, false, false, false, false },
                // outlineLevel
                    new int[] { 1, 2, 3, 3, 3 }
                    );


            // Level 1 collapsed, others expanded, should only have 4 rows, all hidden: 
            verifyGroupCollapsed(
                // level1, level2, level3
                    true, false, false,
                // collapsed:
                    new Boolean[] { false, false, false, false, false },
                // hidden:
                    new bool[] { true, true, true, true, true },
                // outlineLevel
                    new int[] { 1, 2, 3, 3, 3 }
                    );

            // Level 1 and 2 collapsed, Level 3 expanded, 
            verifyGroupCollapsed(
                // level1, level2, level3
                    true, true, false,
                // collapsed:
                    new Boolean[] { false, false, false, false, true, false },
                // hidden:
                    new bool[] { true, true, true, true, true, false },
                // outlineLevel
                    new int[] { 1, 2, 3, 3, 3, 0 }
                    );

            // Level 1 collapsed, Level 2 expanded, Level 3 collapsed 
            verifyGroupCollapsed(
                // level1, level2, level3
                    true, false, true,
                // collapsed:
                    new Boolean[] { false, false, false, false, false, true },
                // hidden:
                    new bool[] { true, true, true, true, true, false },
                // outlineLevel
                    new int[] { 1, 2, 3, 3, 3, 0 }
                    );

            // Level 2 collapsed, others expanded:
            verifyGroupCollapsed(
                // level1, level2, level3
                    false, true, false,
                // collapsed:
                    new Boolean[] { false, false, false, false, false, false },
                // hidden:
                    new bool[] { false, true, true, true, true, false },
                // outlineLevel
                    new int[] { 1, 2, 3, 3, 3, 0 }
                    );

            // Level 3 collapsed, others expanded 
            verifyGroupCollapsed(
                // level1, level2, level3
                    false, false, true,
                // collapsed:
                    new Boolean[] { false, false, false, false, false, true },
                // hidden:
                    new bool[] { false, false, true, true, true, false },
                // outlineLevel
                    new int[] { 1, 2, 3, 3, 3, 0 }
                    );

            // All collapsed 
            verifyGroupCollapsed(
                // level1, level2, level3
                    true, true, true,
                // collapsed:
                    new Boolean[] { false, false, false, false, true, true },
                // hidden:
                    new bool[] { true, true, true, true, true, false },
                // outlineLevel
                    new int[] { 1, 2, 3, 3, 3, 0 }
                    );
        }


        private void verifyGroupCollapsed(bool level1, bool level2, bool level3,
                Boolean[] collapsed, bool[] hidden, int[] outlineLevel)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet123");

            for (int i = 0; i < 4; i++)
            {
                sheet.CreateRow(i);
            }

            sheet.GroupRow(0, 4);
            sheet.GroupRow(1, 4);
            sheet.GroupRow(2, 4);

            sheet.SetRowGroupCollapsed(0, level1);
            sheet.SetRowGroupCollapsed(1, level2);
            sheet.SetRowGroupCollapsed(2, level3);

            CheckWorkbookGrouping(wb, collapsed, hidden, outlineLevel);
        }

        [Test]
        public void Test55640_VerifyCasesSpec()
        {
            // NOTE: This is currently based on current behavior of POI, somehow
            // what POI returns in the calls to collapsed/hidden is not fully matching 
            // the examples in the spec or I did not fully understand how POI stores the data internally...

            // all expanded
            verifyGroupCollapsedSpec(
                // level3, level2, level1
                    false, false, false,
                // collapsed:
                    new Boolean[] { false, false, false, false },
                // hidden:
                    new bool[] { false, false, false, false },
                // outlineLevel
                    new int[] { 3, 3, 2, 1 }
                    );


            verifyGroupCollapsedSpec(
                // level3, level2, level1
                    false, false, true,
                // collapsed:
                    new Boolean[] { false, false, false, true },
                // hidden:
                    new bool[] { true, true, true, false },
                // outlineLevel
                    new int[] { 3, 3, 2, 1 }
                    );

            verifyGroupCollapsedSpec(
                // level3, level2, level1
                    false, true, false,
                // collapsed:
                    new Boolean[] { false, false, true, false },
                // hidden:
                    new bool[] { true, true, true, false },
                // outlineLevel
                    new int[] { 3, 3, 2, 1 }
                    );

            verifyGroupCollapsedSpec(
                // level3, level2, level1
                    false, true, true,
                // collapsed:
                    new Boolean[] { false, false, true, true },
                // hidden:
                    new bool[] { true, true, true, false },
                // outlineLevel
                    new int[] { 3, 3, 2, 1 }
                    );
        }

        private void verifyGroupCollapsedSpec(bool level1, bool level2, bool level3,
                Boolean[] collapsed, bool[] hidden, int[] outlineLevel)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet123");

            for (int i = 5; i < 9; i++)
            {
                sheet.CreateRow(i);
            }

            sheet.GroupRow(5, 6);
            sheet.GroupRow(5, 7);
            sheet.GroupRow(5, 8);

            sheet.SetRowGroupCollapsed(6, level1);
            sheet.SetRowGroupCollapsed(7, level2);
            sheet.SetRowGroupCollapsed(8, level3);

            CheckWorkbookGrouping(wb, collapsed, hidden, outlineLevel);
        }

        private void CheckWorkbookGrouping(IWorkbook wb, bool[] collapsed, bool[] hidden, int[] outlineLevel)
        {
            printWorkbook(wb);
            ISheet sheet = wb.GetSheetAt(0);

            Assert.AreEqual(collapsed.Length, hidden.Length);
            Assert.AreEqual(collapsed.Length, outlineLevel.Length);
            Assert.AreEqual(collapsed.Length, sheet.LastRowNum - sheet.FirstRowNum + 1,
                "Expected " + collapsed.Length + " rows with collapsed state, but had " + (sheet.LastRowNum - sheet.FirstRowNum + 1) + " rows ("
                    + sheet.FirstRowNum + "-" + sheet.LastRowNum + ")");
            for (int i = sheet.FirstRowNum; i < sheet.LastRowNum; i++)
            {
                if (collapsed[i - sheet.FirstRowNum] == null)
                {
                    continue;
                }
                XSSFRow row = (XSSFRow)sheet.GetRow(i);
                Assert.IsNotNull(row, "Could not read row " + i);
                Assert.IsNotNull(row.GetCTRow(), "Could not read row " + i);
                Assert.AreEqual(collapsed[i - sheet.FirstRowNum], row.GetCTRow().collapsed, "Row: " + i + ": collapsed");
                Assert.AreEqual(hidden[i - sheet.FirstRowNum], row.GetCTRow().hidden, "Row: " + i + ": hidden");

                Assert.AreEqual(outlineLevel[i - sheet.FirstRowNum], row.GetCTRow().outlineLevel, "Row: " + i + ": level");
            }

            WriteToFile(wb);
        }
        private void CheckWorkbookGrouping(IWorkbook wb, bool?[] collapsed, bool[] hidden, int[] outlineLevel)
        {
            printWorkbook(wb);
            ISheet sheet = wb.GetSheetAt(0);

            Assert.AreEqual(collapsed.Length, hidden.Length);
            Assert.AreEqual(collapsed.Length, outlineLevel.Length);
            Assert.AreEqual(collapsed.Length, sheet.LastRowNum - sheet.FirstRowNum + 1,
                "Expected " + collapsed.Length + " rows with collapsed state, but had " + (sheet.LastRowNum - sheet.FirstRowNum + 1) + " rows ("
                    + sheet.FirstRowNum + "-" + sheet.LastRowNum + ")");
            for (int i = sheet.FirstRowNum; i < sheet.LastRowNum; i++)
            {
                if (collapsed[i - sheet.FirstRowNum] == null)
                {
                    continue;
                }
                XSSFRow row = (XSSFRow)sheet.GetRow(i);
                Assert.IsNotNull(row, "Could not read row " + i);
                Assert.IsNotNull(row.GetCTRow(), "Could not read row " + i);
                Assert.AreEqual(collapsed[i - sheet.FirstRowNum], row.GetCTRow().collapsed, "Row: " + i + ": collapsed");
                Assert.AreEqual(hidden[i - sheet.FirstRowNum], row.GetCTRow().hidden, "Row: " + i + ": hidden");

                Assert.AreEqual(outlineLevel[i - sheet.FirstRowNum], row.GetCTRow().outlineLevel, "Row: " + i + ": level");
            }

            WriteToFile(wb);
        }

        [Test]
        public void Test55640working()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("sheet123");

            sheet.GroupRow(1, 4);
            sheet.GroupRow(2, 5);
            sheet.GroupRow(3, 6);

            sheet.SetRowGroupCollapsed(1, true);
            sheet.SetRowGroupCollapsed(2, false);
            sheet.SetRowGroupCollapsed(3, false);

            WriteToFile(wb);
        }

        // just used for printing out contents of spreadsheets
        [Test]
        public void Test55640printSample()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("55640.xlsx");
            printWorkbook(wb);

            wb = XSSFTestDataSamples.OpenSampleWorkbook("GroupTest.xlsx");
            printWorkbook(wb);
        }

        private void printWorkbook(IWorkbook wb)
        {
            // disable all output for now...
            //        ISheet sheet = wb.GetSheetAt(0);
            //        
            //        for(Iterator<Row> it = sheet.RowIterator();it.HasNext();) {
            //            XSSFRow row = (XSSFRow) it.Next();
            //            bool collapsed = row.CTRow.Collapsed;
            //            bool hidden = row.CTRow.Hidden;
            //            short level = row.CTRow.OutlineLevel;
            //            
            //            System.out.Println("Row: " + row.RowNum + ": Level: " + level + " Collapsed: " + collapsed + " Hidden: " + hidden);
            //        }
        }

        [Test]
        public void TestGroupingTest()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("GroupTest.xlsx");

            Assert.AreEqual(31, wb.GetSheetAt(0).LastRowNum);

            // NOTE: This is currently based on current behavior of POI, somehow
            // what POI returns in the calls to collapsed/hidden is not fully matching 
            // the examples in the spec or I did not fully understand how POI stores the data internally...
            CheckWorkbookGrouping(wb,
                    new bool?[] {
                    // 0-4
                    false, false, false, false, false, null, null, 
                    // 7-11
                    false, false, true, true, true, null, null, 
                    // 14-18
                    false, false, true, false, false, null,
                    // 20-24
                    false, false, true, true, false, null, null, 
                    // 27-31
                    false, false, false, true, false },
                    new bool[] { 
                    // 0-4
                    false, false, false, false, false, false, false,  
                    // 7-11
                    true,  true, true, true, false, false, false, 
                    // 14-18
                    true, true, false, false, false, false,  
                    // 20-24
                    true, true, true, false, false, false, false, 
                    // 27-31
                    true, true, true, true, false },
                // outlineLevel
                    new int[] { 
                    // 0-4
                    3, 3, 2, 1, 0, 0, 0,
                    // 7-11
                    3, 3, 2, 1, 0, 0, 0,
                    // 14-18
                    3, 3, 2, 1, 0, 0,
                    // 20-24
                    3, 3, 2, 1, 0, 0, 0, 
                    // 27-31
                    3, 3, 2, 1, 0,
                }
                    );
            /*
    Row: 0: Level: 3 Collapsed: false Hidden: false
    Row: 1: Level: 3 Collapsed: false Hidden: false
    Row: 2: Level: 2 Collapsed: false Hidden: false
    Row: 3: Level: 1 Collapsed: false Hidden: false
    Row: 4: Level: 0 Collapsed: false Hidden: false
    Row: 7: Level: 3 Collapsed: false Hidden: true
    Row: 8: Level: 3 Collapsed: false Hidden: true
    Row: 9: Level: 2 Collapsed: true Hidden: true
    Row: 10: Level: 1 Collapsed: true Hidden: true
    Row: 11: Level: 0 Collapsed: true Hidden: false
    Row: 14: Level: 3 Collapsed: false Hidden: true
    Row: 15: Level: 3 Collapsed: false Hidden: true
    Row: 16: Level: 2 Collapsed: true Hidden: false
    Row: 17: Level: 1 Collapsed: false Hidden: false
    Row: 18: Level: 0 Collapsed: false Hidden: false
    Row: 20: Level: 3 Collapsed: false Hidden: true
    Row: 21: Level: 3 Collapsed: false Hidden: true
    Row: 22: Level: 2 Collapsed: true Hidden: true
    Row: 23: Level: 1 Collapsed: true Hidden: false
    Row: 24: Level: 0 Collapsed: false Hidden: false
    Row: 27: Level: 3 Collapsed: false Hidden: true
    Row: 28: Level: 3 Collapsed: false Hidden: true
    Row: 29: Level: 2 Collapsed: false Hidden: true
    Row: 30: Level: 1 Collapsed: true Hidden: true
    Row: 31: Level: 0 Collapsed: true Hidden: false
             */
        }
    }
}