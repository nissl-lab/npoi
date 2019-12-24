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
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.HSSF.UserModel;
    using TestCases.HSSF.UserModel;
    using NPOI.Util;
    using System.IO;
    using NPOI.SS.Util;
    using NPOI.HSSF.Model;
    using System.Text;

    /**
     * 
     */
    [TestFixture]
    public class TestRowRecordsAggregate
    {
        [Test]
        public void TestRowGet()
        {
            RowRecordsAggregate rra = new RowRecordsAggregate();
            RowRecord rr = new RowRecord(4);
            rra.InsertRow(rr);
            rra.InsertRow(new RowRecord(1));

            RowRecord rr1 = rra.GetRow(4);

            Assert.IsNotNull(rr1);
            Assert.AreEqual(4, rr1.RowNumber, "Row number is1 1");
            Assert.IsTrue(rr1 == rr, "Row record retrieved is1 identical");
        }
        /**
	 * Prior to Aug 2008, POI would re-serialize spreadsheets with {@link ArrayRecord}s or
	 * {@link TableRecord}s with those records out of order.  Similar to
	 * {@link SharedFormulaRecord}s, these records should appear immediately after the first
	 * {@link FormulaRecord}s that they apply to (and only once).<br/>
	 */
        [Test]
        public void TestArraysAndTables()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("testArraysAndTables.xls");
            Record[] sheetRecs = RecordInspector.GetRecords(wb.GetSheetAt(0), 0);

            int countArrayFormulas = verifySharedValues(sheetRecs, typeof(ArrayRecord));
            Assert.AreEqual(5, countArrayFormulas);
            int countTableFormulas = verifySharedValues(sheetRecs, typeof(TableRecord));
            Assert.AreEqual(3, countTableFormulas);

            // Note - SharedFormulaRecords are currently not re-serialized by POI (each is extracted
            // into many non-shared formulas), but if they ever were, the same rules would apply.
            int countSharedFormulas = verifySharedValues(sheetRecs, typeof(SharedFormulaRecord));
            Assert.AreEqual(0, countSharedFormulas);


            //if (false) { // set true to observe re-serialized file
            //    File f = new File(System.getProperty("java.io.tmpdir") + "/testArraysAndTables-out.xls");
            //    try {
            //        OutputStream os = new FileOutputStream(f);
            //        wb.write(os);
            //        os.close();
            //    } catch (IOException e) {
            //        throw new RuntimeException(e);
            //    }
            //    System.Console.WriteLine("Output file to " + f.getAbsolutePath());
            //}

            wb.Close();
        }

        private static int verifySharedValues(Record[] recs, Type shfClass)
        {

            int result = 0;
            for (int i = 0; i < recs.Length; i++)
            {
                Record rec = recs[i];
                if (rec.GetType() == shfClass)
                {
                    result++;
                    Record prevRec = recs[i - 1];
                    if (!(prevRec is FormulaRecord))
                    {
                        Assert.Fail("Bad record order at index "
                                + i + ": Formula record expected but got ("
                                + prevRec.GetType().Name + ")");
                    }
                    verifySharedFormula((FormulaRecord)prevRec, rec);
                }
            }
            return result;
        }

        private static void verifySharedFormula(FormulaRecord firstFormula, Record rec)
        {
            CellRangeAddress8Bit range = ((SharedValueRecordBase)rec).Range;
            Assert.AreEqual(range.FirstRow, firstFormula.Row);
            Assert.AreEqual(range.FirstColumn, firstFormula.Column);
        }

        /**
         * This problem was noted as the overt symptom of bug 46280.  The logic for skipping {@link
         * UnknownRecord}s in the constructor {@link RowRecordsAggregate} did not allow for the
         * possibility of tailing {@link ContinueRecord}s.<br/>
         * The functionality change being tested here is actually not critical to the overall fix
         * for bug 46280, since the fix involved making sure the that offending <i>PivotTable</i>
         * records do not get into {@link RowRecordsAggregate}.<br/>
         * This fix in {@link RowRecordsAggregate} was implemented anyway since any {@link
         * UnknownRecord} has the potential of being 'continued'.
         */
        [Test]
        public void TestUnknownContinue_bug46280()
        {
            byte[] dummtydata = Encoding.GetEncoding(1252).GetBytes("dummydata");
            byte[] moredummydata = Encoding.GetEncoding(1252).GetBytes("moredummydata");
            Record[] inRecs = {
			new RowRecord(0),
			new NumberRecord(),
            new UnknownRecord(0x5555,dummtydata),
            new ContinueRecord(moredummydata)
			//new UnknownRecord(0x5555, "dummydata".getBytes()),
			//new ContinueRecord("moredummydata".getBytes()),
		};
            RecordStream rs = new RecordStream(Arrays.AsList(inRecs), 0);
            RowRecordsAggregate rra;
            try
            {
                rra = new RowRecordsAggregate(rs, SharedValueManager.CreateEmpty());
            }
            catch (RuntimeException e)
            {
                if (e.Message.StartsWith("Unexpected record type"))
                {
                    Assert.Fail("Identified bug 46280a");
                }
                throw e;
            }
            RecordInspector.RecordCollector rv = new RecordInspector.RecordCollector();
            rra.VisitContainedRecords(rv);
            Record[] outRecs = rv.Records;
            Assert.AreEqual(5, outRecs.Length);
        }
    }
}