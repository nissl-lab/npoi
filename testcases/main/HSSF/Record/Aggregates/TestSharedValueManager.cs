/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.Reflection;
using NUnit.Framework;
using NPOI.HSSF.Record;
using NPOI.HSSF.Record.Aggregates;
using NPOI.HSSF.UserModel;
using NPOI.Util;
using TestCases.HSSF.UserModel;
namespace TestCases.HSSF.Record.Aggregates
{

    /**
     * Tests for {@link SharedValueManager}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestSharedValueManager
    {

        /**
         * This Excel workbook contains two sheets that each have a pair of overlapping shared formula
         * ranges.  The first sheet has one row and one column shared formula ranges which intersect.
         * The second sheet has two column shared formula ranges - one contained within the other.
         * These shared formula ranges were created by fill-dragging a single cell formula across the
         * desired region.  The larger shared formula ranges were placed first.<br/>
         *
         * There are probably many ways to produce similar effects, but it should be noted that Excel
         * is quite temperamental in this regard.  Slight variations in technique can cause the shared
         * formulas to spill out into plain formula records (which would make these tests pointless).
         *
         */
        private static String SAMPLE_FILE_NAME = "overlapSharedFormula.xls";
        /**
         * Some of these bugs are intermittent, and the test author couldn't think of a way to write
         * test code to hit them bug deterministically. The reason for the unpredictability is that
         * the bugs depended on the {@link SharedFormulaRecord}s being searched in a particular order.
         * At the time of writing of the test, the order was being determined by the call to {@link
         * Collection#toArray(Object[])} on {@link HashMap#values()} where the items in the map were
         * using default {@link Object#hashCode()}<br/>
         */
        private static int MAX_ATTEMPTS = 5;

        /**
         * This bug happened when there were two or more shared formula ranges that overlapped.  POI
         * would sometimes associate formulas in the overlapping region with the wrong shared formula
         */
        [Test]
        public void TestPartiallyOverlappingRanges()
        {
            NPOI.HSSF.Record.Record[] records;

            int attempt = 1;
            do
            {
                HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook(SAMPLE_FILE_NAME);

                HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(0);
                RecordInspector.GetRecords(sheet, 0);
                Assert.AreEqual("1+1", sheet.GetRow(2).GetCell(0).CellFormula);
                if ("1+1".Equals(sheet.GetRow(3).GetCell(0).CellFormula))
                {
                    throw new AssertionException("Identified bug - wrong shared formula record chosen"
                            + " (attempt " + attempt + ")");
                }
                Assert.AreEqual("2+2", sheet.GetRow(3).GetCell(0).CellFormula);
                records = RecordInspector.GetRecords(sheet, 0);
            } while (attempt++ < MAX_ATTEMPTS);

            int count = 0;
            for (int i = 0; i < records.Length; i++)
            {
                if (records[i] is SharedFormulaRecord)
                {
                    count++;
                }
            }
            Assert.AreEqual(2, count);
        }

        /**
         * This bug occurs for similar reasons to the bug in {@link #testPartiallyOverlappingRanges()}
         * but the symptoms are much uglier - serialization fails with {@link NullPointerException}.<br/>
         */
        [Test]
        public void TestCompletelyOverlappedRanges()
        {
            NPOI.HSSF.Record.Record[] records;

            int attempt = 1;
            do
            {
                HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook(SAMPLE_FILE_NAME);

                HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(1);
                try
                {
                    records = RecordInspector.GetRecords(sheet, 0);
                }
                catch (NullReferenceException)
                {
                    throw new AssertionException("Identified bug " +
                            "- cannot reserialize completely overlapped shared formula"
                            + " (attempt " + attempt + ")");
                }
            } while (attempt++ < MAX_ATTEMPTS);

            int count = 0;
            for (int i = 0; i < records.Length; i++)
            {
                if (records[i] is SharedFormulaRecord)
                {
                    count++;
                }
            }
            Assert.AreEqual(2, count);
        }

        /**
         * Tests fix for a bug in the way shared formula cells are associated with shared formula
         * records.  Prior to this fix, POI would attempt to use the upper left corner of the
         * shared formula range as the locator cell.  The correct cell to use is the 'first cell'
         * in the shared formula group which is not always the top left cell.  This is possible
         * because shared formula groups may be sparse and may overlap.<br/>
         *
         * Two existing sample files (15228.xls and ex45046-21984.xls) had similar issues.
         * These were not explored fully, but seem to be fixed now.
         */
        [Test]
        public void TestRecalculateFormulas47747()
        {

            /*
             * ex47747-sharedFormula.xls is a heavily cut-down version of the spreadsheet from
             * the attachment (id=24176) in Bugzilla 47747.  This was done to make the sample
             * file smaller, which hopefully allows the special data encoding condition to be
             * seen more easily.  Care must be taken when modifying this file since the
             * special conditions are easily destroyed (which would make this test useless).
             * It seems that removing the worksheet protection has made this more so - if the
             * current file is re-saved in Excel(2007) the bug condition disappears.
             *
             *
             * Using BiffViewer, one can see that there are two shared formula groups representing
             * the essentially same formula over ~20 cells.  The shared group ranges overlap and
             * are A12:Q20 and A20:Q27.  The locator cell ('first cell') for the second group is
             * Q20 which is not the top left cell of the enclosing range.  It is this specific
             * condition which caused the bug to occur
             */
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex47747-sharedFormula.xls");

            // pick out a cell from within the second shared formula group
            HSSFCell cell = (HSSFCell)wb.GetSheetAt(0).GetRow(23).GetCell(0);
            String formulaText;
            try
            {
                formulaText = cell.CellFormula;
                // succeeds if the formula record has been associated
                // with the second shared formula group
            }
            catch (RuntimeException e)
            {
                // bug occurs if the formula record has been associated
                // with the first shared formula group
                if ("Shared Formula Conversion: Coding Error".Equals(e.Message))
                {
                    throw new AssertionException("Identified bug 47747");
                }
                throw e;
            }
            Assert.AreEqual("$AF24*A$7", formulaText);
        }

        /**
         * Convenience test method for digging the {@link SharedValueManager} out of a
         * {@link RowRecordsAggregate}.
         */
        public static SharedValueManager ExtractFromRRA(RowRecordsAggregate rra)
        {
            FieldInfo f;
            try
            {
                f = typeof(RowRecordsAggregate).GetField("_sharedValueManager", BindingFlags.NonPublic | BindingFlags.Instance);
                //typeof(RowRecordsAggregate).("_sharedValueManager");
            }
            catch (NotSupportedException e)
            {
                throw new RuntimeException(e);
            }
            //f.s
            //f.setAccessible(true);
            try
            {
                return (SharedValueManager)f.GetValue(rra);
            }
            catch (ArgumentException e)
            {
                throw new RuntimeException(e);
            }
            catch (FieldAccessException e)
            {
                throw new RuntimeException(e);
            }
        }
        [Test]
        public void TestBug52527()
        {
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("52527.xls");
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);

            Assert.AreEqual("IF(H3,LINEST(N9:N14,K9:M14,FALSE),LINEST(N8:N14,K8:M14,FALSE))",
                    wb1.GetSheetAt(0).GetRow(4).GetCell(11).CellFormula);
            Assert.AreEqual("IF(H3,LINEST(N9:N14,K9:M14,FALSE),LINEST(N8:N14,K8:M14,FALSE))",
                    wb2.GetSheetAt(0).GetRow(4).GetCell(11).CellFormula);

            Assert.AreEqual("1/SQRT(J9)",
                    wb1.GetSheetAt(0).GetRow(8).GetCell(10).CellFormula);
            Assert.AreEqual("1/SQRT(J9)",
                    wb2.GetSheetAt(0).GetRow(8).GetCell(10).CellFormula);

            Assert.AreEqual("1/SQRT(J26)",
                    wb1.GetSheetAt(0).GetRow(25).GetCell(10).CellFormula);
            Assert.AreEqual("1/SQRT(J26)",
                    wb2.GetSheetAt(0).GetRow(25).GetCell(10).CellFormula);
        }
    }
}