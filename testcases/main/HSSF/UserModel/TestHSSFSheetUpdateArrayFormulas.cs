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
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using TestCases.HSSF.Record.Aggregates;

    /**
     * Test array formulas in HSSF
     *
     * @author Yegor Kozlov
     * @author Josh Micich
     */
    [TestFixture]
    public class TestHSSFSheetUpdateArrayFormulas
    {

        // Test methods common with XSSF are in superclass
        // Local methods here test HSSF-specific details of updating array formulas

        [Test]
        public void TestHSSFSetArrayFormula_SingleCell()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            CellRangeAddress range = new CellRangeAddress(2, 2, 2, 2);
            ICell[] cells = sheet.SetArrayFormula("SUM(C11:C12*D11:D12)", range).FlattenedCells;
            Assert.AreEqual(1, cells.Length);

            // sheet.SetArrayFormula Creates rows and cells for the designated range
            Assert.IsNotNull(sheet.GetRow(2));
            ICell cell = sheet.GetRow(2).GetCell(2);
            Assert.IsNotNull(cell);

            Assert.IsTrue(cell.IsPartOfArrayFormulaGroup);
            //retrieve the range and check it is the same
            Assert.AreEqual(range.FormatAsString(), cell.ArrayFormulaRange.FormatAsString());

            FormulaRecordAggregate agg = (FormulaRecordAggregate)(((HSSFCell)cell).CellValueRecord);
            Assert.AreEqual(range.FormatAsString(), agg.GetArrayFormulaRange().FormatAsString());
            Assert.IsTrue(agg.IsPartOfArrayFormula);
            workbook.Close();
        }

        /**
         * Makes sure the internal state of HSSFSheet is consistent After removing array formulas
         */
        [Test]
        public void TestAddRemoveArrayFormulas_recordUpdates()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet("Sheet1");

            ICellRange<ICell> cr = s.SetArrayFormula("123", CellRangeAddress.ValueOf("B5:C6"));
            Record[] recs;
            int ix;
            recs = RecordInspector.GetRecords(s, 0);
            ix = FindRecordOfType(recs, typeof(ArrayRecord), 0);
            ConfirmRecordClass(recs, ix - 1, typeof(FormulaRecord));
            ConfirmRecordClass(recs, ix + 1, typeof(FormulaRecord));
            ConfirmRecordClass(recs, ix + 2, typeof(FormulaRecord));
            ConfirmRecordClass(recs, ix + 3, typeof(FormulaRecord));
            // just one array record
            Assert.IsTrue(FindRecordOfType(recs, typeof(ArrayRecord), ix + 1) < 0);

            s.RemoveArrayFormula(cr.TopLeftCell);

            // Make sure the array formula has been Removed properly

            recs = RecordInspector.GetRecords(s, 0);
            Assert.IsTrue(FindRecordOfType(recs, typeof(ArrayRecord), 0) < 0);
            Assert.IsTrue(FindRecordOfType(recs, typeof(FormulaRecord), 0) < 0);
            RowRecordsAggregate rra = ((HSSFSheet)s).Sheet.RowsAggregate;
            SharedValueManager svm = TestSharedValueManager.ExtractFromRRA(rra);
            if (svm.GetArrayRecord(4, 1) != null)
            {
                Assert.Fail("Array record was not cleaned up properly.");
            }
            wb.Close();
        }

        private static void ConfirmRecordClass(Record[] recs, int index, Type cls)
        {
            if (recs.Length <= index)
            {
                Assert.Fail("Expected (" + cls.Name + ") at index "
                        + index + " but array length is " + recs.Length + ".");
            }
            Assert.AreEqual(cls, recs[index].GetType());
        }
        /**
         * @return <tt>-1<tt> if not found
         */
        private static int FindRecordOfType(Record[] recs, Type type, int fromIndex)
        {
            for (int i = fromIndex; i < recs.Length; i++)
            {
                if (type.IsInstanceOfType(recs[i]))
                {
                    return i;
                }
            }
            return -1;
        }
    }
}
