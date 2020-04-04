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

namespace TestCases.HSSF.Model
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.PivotTable;
    using NPOI.Util;
    using NUnit.Framework;
    using NPOI.HSSF.Model;

    /**
     * Tests for {@link RowBlocksReader}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRowBlocksReader
    {
        [Test]
        public void TestAbnormalPivotTableRecords_bug46280()
        {
            int SXVIEW_SID = ViewDefinitionRecord.sid;
            Record[] inRecs = {
			new RowRecord(0),
			new NumberRecord(),
			// normally MSODRAWING(0x00EC) would come here before SXVIEW
			new UnknownRecord(SXVIEW_SID, System.Text.Encoding.UTF8.GetBytes("dummydata (SXVIEW: View DefInition)")),
			new WindowTwoRecord(),
		};
            RecordStream rs = new RecordStream(Arrays.AsList(inRecs), 0);
            RowBlocksReader rbr = new RowBlocksReader(rs);
            if (rs.PeekNextClass() == typeof(WindowTwoRecord))
            {
                // Should have stopped at the SXVIEW record
                Assert.Fail("Identified bug 46280b");
            }
            RecordStream rbStream = rbr.PlainRecordStream;
            Assert.AreEqual(inRecs[0], rbStream.GetNext());
            Assert.AreEqual(inRecs[1], rbStream.GetNext());
            Assert.IsFalse(rbStream.HasNext());
            Assert.IsTrue(rs.HasNext());
            Assert.AreEqual(inRecs[2], rs.GetNext());
            Assert.AreEqual(inRecs[3], rs.GetNext());
        }
    }
}


