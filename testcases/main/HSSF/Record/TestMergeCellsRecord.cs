
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

namespace TestCases.HSSF.Record
{

    using System;
    using NPOI.HSSF.Record;
    using NPOI.SS.Util;
    using NUnit.Framework;
    using NPOI.HSSF.Record.Aggregates;
    using System.Collections;
    using NPOI.HSSF.Model;

    /**
     * Make sure the merge cells record behaves
     * @author Danny Mui (dmui at apache dot org)
     *
     */
    [TestFixture]
    public class TestMergeCellsRecord
    {

        /**
         * Make sure when a clone is1 called, we actually clone it.
         * @
         */
        [Test]
        public void TestCloneReferences()
        {
            CellRangeAddress[] cras = { new CellRangeAddress(0, 1, 0, 2), };
            MergeCellsRecord merge = new MergeCellsRecord(cras, 0, cras.Length);
            MergeCellsRecord clone = (MergeCellsRecord)merge.Clone();

            Assert.AreNotSame(merge, clone, "Merged and cloned objects are the same");

            CellRangeAddress mergeRegion = merge.GetAreaAt(0);
            CellRangeAddress cloneRegion = clone.GetAreaAt(0);
            Assert.AreNotSame(mergeRegion, cloneRegion,
                "Should not point to same objects when cloning");
            Assert.AreEqual(mergeRegion.FirstRow, cloneRegion.FirstRow,
                "New Clone Row From doesnt match");
            Assert.AreEqual(mergeRegion.LastRow, cloneRegion.LastRow,
                "New Clone Row To doesnt match");
            Assert.AreEqual(mergeRegion.FirstColumn, cloneRegion.FirstColumn,
                "New Clone Col From doesnt match");
            Assert.AreEqual(mergeRegion.LastColumn, cloneRegion.LastColumn,
                "New Clone Col To doesnt match");

            Assert.IsFalse(merge.GetAreaAt(0) == clone.GetAreaAt(0));
        }
        [Test]
        public void testMCTable_bug46009()
        {
            MergedCellsTable mct = new MergedCellsTable();
            ArrayList recList = new ArrayList();
            CellRangeAddress[] cras = new CellRangeAddress[] {
				new CellRangeAddress(0, 0, 0, 3), 
		};
            recList.Add(new MergeCellsRecord(cras, 0, 1));
            RecordStream rs = new RecordStream(recList, 0);
            mct.Read(rs);
            try
            {
                mct.VisitContainedRecords(dummyRecordVisitor);
            }
            catch (Exception)
            {
                throw new AssertionException("Identified bug 46009");
            }
        }
        DummyRecordVisitor dummyRecordVisitor = new DummyRecordVisitor();
        private class DummyRecordVisitor : RecordVisitor
        {
            public void VisitRecord(Record r)
            {
                // do nothing
            }
        }
    }
}