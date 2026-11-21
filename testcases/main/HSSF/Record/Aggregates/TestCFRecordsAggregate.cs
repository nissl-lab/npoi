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
    using System.IO;
    using System.Collections;

    using NUnit.Framework;using NUnit.Framework.Legacy;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Util;
    using NPOI.SS.Util;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;
    using NPOI.HSSF.Model;
    using NPOI.Util;

    /**
     * Tests the serialization and deserialization of the CFRecordsAggregate
     * class works correctly.  
     *
     * @author Dmitriy Kumshayev 
     */
    [TestFixture]
    public class TestCFRecordsAggregate
    {

        [Test]
        public void TestCFRecordsAggregate1()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet();
            List<Record> recs = new List<Record>();
            CFHeaderBase header = new CFHeaderRecord();
            CFRuleBase rule1 = CFRuleRecord.Create(sheet, "7");
            CFRuleBase rule2 = CFRuleRecord.Create(sheet, (byte)ComparisonOperator.Between, "2", "5");
            CFRuleBase rule3 = CFRuleRecord.Create(sheet, (byte)ComparisonOperator.GreaterThanOrEqual, "100", null);
            header.NumberOfConditionalFormats = (3);
            CellRangeAddress[] cellRanges = {
				new CellRangeAddress(0,1,0,0),
				new CellRangeAddress(0,1,2,2),
		    };
            header.CellRanges = (cellRanges);
            recs.Add(header);
            recs.Add(rule1);
            recs.Add(rule2);
            recs.Add(rule3);
            CFRecordsAggregate record = CFRecordsAggregate.CreateCFAggregate(new RecordStream(recs, 0));

            // Serialize
            byte [] serializedRecord = new byte[record.RecordSize];
		    record.Serialize(0, serializedRecord);
            ByteArrayInputStream in1 = new ByteArrayInputStream(serializedRecord);

            //Parse
            recs = RecordFactory.CreateRecords(in1);

            // Verify
            ClassicAssert.IsNotNull(recs);
            ClassicAssert.AreEqual(4, recs.Count);

            header = (CFHeaderRecord)recs[0];
            rule1 = (CFRuleRecord)recs[1];
            ClassicAssert.IsNotNull(rule1);
            rule2 = (CFRuleRecord)recs[2];
            ClassicAssert.IsNotNull(rule2);
            rule3 = (CFRuleRecord)recs[3];
            ClassicAssert.IsNotNull(rule3);
            cellRanges = header.CellRanges;

            ClassicAssert.AreEqual(2, cellRanges.Length);
            ClassicAssert.AreEqual(3, header.NumberOfConditionalFormats);
            ClassicAssert.IsFalse(header.NeedRecalculation);

            record = CFRecordsAggregate.CreateCFAggregate(new RecordStream(recs, 0));

            record = record.CloneCFAggregate();

            ClassicAssert.IsNotNull(record.Header);
            ClassicAssert.AreEqual(3, record.NumberOfRules);

            header = record.Header;
            rule1 = record.GetRule(0);
            ClassicAssert.IsNotNull(rule1);
            rule2 = record.GetRule(1);
            ClassicAssert.IsNotNull(rule2);
            rule3 = record.GetRule(2);
            ClassicAssert.IsNotNull(rule3);
            cellRanges = header.CellRanges;

            ClassicAssert.AreEqual(2, cellRanges.Length);
            ClassicAssert.AreEqual(3, header.NumberOfConditionalFormats);
            ClassicAssert.IsFalse(header.NeedRecalculation);
        }
        /**
         * Make sure that the CF Header record is properly updated with the number of rules
         */
        [Test]
        public void TestNRules()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet();
            CellRangeAddress[] cellRanges = {
				new CellRangeAddress(0,1,0,0),
				new CellRangeAddress(0,1,2,2),
		    };
            CFRuleRecord[] rules = {
			CFRuleRecord.Create(sheet, "7"),
			CFRuleRecord.Create(sheet, (byte)ComparisonOperator.Between, "2", "5"),
		    };
            CFRecordsAggregate agg = new CFRecordsAggregate(cellRanges, rules);
            byte[] serializedRecord = new byte[agg.RecordSize];
            agg.Serialize(0, serializedRecord);

            int nRules = NPOI.Util.LittleEndian.GetUShort(serializedRecord, 4);
            if (nRules == 0)
            {
                throw new AssertionException("Identified bug 45682 b");
            }
            ClassicAssert.AreEqual(rules.Length, nRules);
        }

        [Test]
        public void TestCantMixTypes()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;
            CellRangeAddress[] cellRanges = {
                new CellRangeAddress(0,1,0,0),
                new CellRangeAddress(0,1,2,2),
            };
            CFRuleBase[] rules = {
                CFRuleRecord.Create(sheet, "7"),
                CFRule12Record.Create(sheet, (byte)ComparisonOperator.Between, "2", "5"),
            };
            try
            {
                new CFRecordsAggregate(cellRanges, rules);
                Assert.Fail("Shouldn't be able to mix between types");
            }
            catch (ArgumentException) { }

            rules = new CFRuleBase[] { CFRuleRecord.Create(sheet, "7") };
            CFRecordsAggregate agg = new CFRecordsAggregate(cellRanges, rules);
            ClassicAssert.IsTrue(agg.Header.NeedRecalculation);

            try
            {
                agg.AddRule(CFRule12Record.Create(sheet, "7"));
                Assert.Fail("Shouldn't be able to mix between types");
            }
            catch (ArgumentException) { }
        }

    }
}