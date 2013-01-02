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

namespace TestCases.HSSF.Record.Aggregates
{
    using System;
    using System.Collections;
    using NPOI.HSSF.Record;
    using NUnit.Framework;
    using NPOI.HSSF.Record.Aggregates;

    /**
     * @author Glen Stampoultzis
     */
    [TestFixture]
    public class TestColumnInfoRecordsAggregate
    {
        //ColumnInfoRecordsAggregate columnInfoRecordsAggregate;
        [Test]
        public void TestRecordSize()
        {
            ColumnInfoRecordsAggregate agg = new ColumnInfoRecordsAggregate();
            agg.InsertColumn(CreateColInfo(1, 3));
            agg.InsertColumn(CreateColInfo(4, 7));
            agg.InsertColumn(CreateColInfo(8, 8));
            agg.GroupColumnRange(2, 5, true);
            Assert.AreEqual(4, agg.NumColumns);

            ConfirmSerializedSize(agg);

            agg = new ColumnInfoRecordsAggregate();
            agg.GroupColumnRange(3, 6, true);
            ConfirmSerializedSize(agg);
        }

        private static void ConfirmSerializedSize(RecordBase cirAgg)
        {
            int estimatedSize = cirAgg.RecordSize;
            byte[] buf = new byte[estimatedSize];
            int serializedSize = cirAgg.Serialize(0, buf);
            Assert.AreEqual(estimatedSize, serializedSize);
        }

        private static ColumnInfoRecord CreateColInfo(int firstCol, int lastCol)
        {
            ColumnInfoRecord columnInfoRecord = new ColumnInfoRecord();
            columnInfoRecord.FirstColumn = ((short)firstCol);
            columnInfoRecord.LastColumn = ((short)lastCol);
            return columnInfoRecord;
        }

        private class CIRCollector : RecordVisitor
        {

            private ArrayList _list;
            public CIRCollector()
            {
                _list = new ArrayList();
            }
            public void VisitRecord(Record r)
            {
                _list.Add(r);
            }
            public static ColumnInfoRecord[] GetRecords(ColumnInfoRecordsAggregate agg)
            {
                CIRCollector circ = new CIRCollector();
                agg.VisitContainedRecords(circ);
                ArrayList list = circ._list;
                ColumnInfoRecord[] result = new ColumnInfoRecord[list.Count];
                result = (ColumnInfoRecord[])list.ToArray(typeof(ColumnInfoRecord));
                return result;
            }
        }
        [Test]
        public void TestGroupColumns_bug45639()
        {
            ColumnInfoRecordsAggregate agg = new ColumnInfoRecordsAggregate();
            agg.GroupColumnRange(7, 9, true);
            agg.GroupColumnRange(4, 12, true);
            try
            {
                agg.GroupColumnRange(1, 15, true);
            }
            catch (IndexOutOfRangeException)
            {
                throw new AssertionException("Identified bug 45639");
            }
            ColumnInfoRecord[] cirs = CIRCollector.GetRecords(agg);
            Assert.AreEqual(5, cirs.Length);
            ConfirmCIR(cirs, 0, 1, 3, 1, false, false);
            ConfirmCIR(cirs, 1, 4, 6, 2, false, false);
            ConfirmCIR(cirs, 2, 7, 9, 3, false, false);
            ConfirmCIR(cirs, 3, 10, 12, 2, false, false);
            ConfirmCIR(cirs, 4, 13, 15, 1, false, false);
        }

        /**
         * Check that an inner Group remains hidden
         */
        [Test]
        public void TestHiddenAfterExpanding()
        {
            ColumnInfoRecordsAggregate agg = new ColumnInfoRecordsAggregate();
            agg.GroupColumnRange(1, 15, true);
            agg.GroupColumnRange(4, 12, true);

            ColumnInfoRecord[] cirs;

            // collapse both inner and outer Groups
            agg.CollapseColumn(6);
            agg.CollapseColumn(3);

            cirs = CIRCollector.GetRecords(agg);
            Assert.AreEqual(5, cirs.Length);
            ConfirmCIR(cirs, 0, 1, 3, 1, true, false);
            ConfirmCIR(cirs, 1, 4, 12, 2, true, false);
            ConfirmCIR(cirs, 2, 13, 13, 1, true, true);
            ConfirmCIR(cirs, 3, 14, 15, 1, true, false);
            ConfirmCIR(cirs, 4, 16, 16, 0, false, true);

            // just expand the inner Group
            agg.ExpandColumn(6);

            cirs = CIRCollector.GetRecords(agg);
            Assert.AreEqual(4, cirs.Length);
            if (!cirs[1].IsHidden)
            {
                throw new AssertionException("Inner Group should still be hidden");
            }
            ConfirmCIR(cirs, 0, 1, 3, 1, true, false);
            ConfirmCIR(cirs, 1, 4, 12, 2, true, false);
            ConfirmCIR(cirs, 2, 13, 15, 1, true, false);
            ConfirmCIR(cirs, 3, 16, 16, 0, false, true);
        }
        private static void ConfirmCIR(ColumnInfoRecord[] cirs,
            int ix, int startColIx, int endColIx, int level, bool isHidden, bool isCollapsed)
        {
            ColumnInfoRecord cir = cirs[ix];
            Assert.AreEqual(startColIx, cir.FirstColumn, "startColIx");
            Assert.AreEqual(endColIx, cir.LastColumn, "endColIx");
            Assert.AreEqual(level, cir.OutlineLevel, "level");
            Assert.AreEqual(isHidden, cir.IsHidden, "hidden");
            Assert.AreEqual(isCollapsed, cir.IsCollapsed, "collapsed");
        }
    }
}