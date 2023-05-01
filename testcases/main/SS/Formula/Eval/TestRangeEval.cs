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

namespace TestCases.SS.Formula.Eval
{

    using System;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.HSSF.Util;

    /**
     * Test for unary plus operator Evaluator.
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRangeEval
    {
        [Test]
        public void TestPermutations()
        {

            Confirm("B3", "D7", 2, 6, 1, 3);
            Confirm("B1", "B1", 0, 0, 1, 1);

            Confirm("B7", "D3", 2, 6, 1, 3);
            Confirm("D3", "B7", 2, 6, 1, 3);
            Confirm("D7", "B3", 2, 6, 1, 3);
        }

        private static void Confirm(
            String refA,
            String refB,
            int firstRow,
            int lastRow,
            int firstColumn,
            int lastColumn)
        {
            ValueEval[] args =
            {
                CreateRefEval(refA),
                CreateRefEval(refB),
            };

            ValueEval result = EvalInstances.Range.Evaluate(args, 0, (short)0);
            Assert.IsTrue(result is AreaEval);
            AreaEval ae = (AreaEval)result;
            Assert.AreEqual(firstRow, ae.FirstRow);
            Assert.AreEqual(lastRow, ae.LastRow);
            Assert.AreEqual(firstColumn, ae.FirstColumn);
            Assert.AreEqual(lastColumn, ae.LastColumn);
        }

        private static ValueEval CreateRefEval(String refStr)
        {
            CellReference cr = new CellReference(refStr);
            return new MockRefEval(cr.Row, cr.Col);

        }

        private class MockRefEval : RefEvalBase
        {

            public MockRefEval(int rowIndex, int columnIndex)
                : base(-1, -1, rowIndex, columnIndex)
            {
                ;
            }
            public override ValueEval GetInnerValueEval(int sheetIndex)
            {
                throw new SystemException("not expected to be called during this Test");
            }
            public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx,
                    int relLastColIx)
            {
                AreaI area = new OffsetArea(Row, Column,
                        relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);
                return new MockAreaEval(area);
            }
        }

        private class MockAreaEval : AreaEvalBase
        {

            public MockAreaEval(AreaI ptg)
                : base(ptg)
            {
                ;
            }
            private MockAreaEval(int firstRow, int firstColumn, int lastRow, int lastColumn)
                : base(firstRow, firstColumn, lastRow, lastColumn)
            {
                ;
            }
            public override ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex)
            {
                throw new RuntimeException("not expected to be called during this Test");
            }

            public override ValueEval GetRelativeValue(int sheetIndex, int relativeRowIndex, int relativeColumnIndex)
            {
                throw new RuntimeException("not expected to be called during this test");
            }

            public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx,
                    int relLastColIx)
            {
                AreaI area = new OffsetArea(FirstRow, FirstColumn,
                        relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);

                return new MockAreaEval(area);
            }
            public override TwoDEval GetRow(int rowIndex)
            {
                if (rowIndex >= Height)
                {
                    throw new ArgumentException("Invalid rowIndex " + rowIndex
                            + ".  Allowable range is (0.." + Height + ").");
                }
                return new MockAreaEval(rowIndex, FirstColumn, rowIndex, LastColumn);
            }
            public override TwoDEval GetColumn(int columnIndex)
            {
                if (columnIndex >= Width)
                {
                    throw new ArgumentException("Invalid columnIndex " + columnIndex
                            + ".  Allowable range is (0.." + Width + ").");
                }
                return new MockAreaEval(FirstRow, columnIndex, LastRow, columnIndex);
            }
        }
        [Test]
        public void TestRangeUsingOffsetFunc_bug46948()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            IRow row = wb.CreateSheet("Sheet1").CreateRow(0);
            ICell cellA1 = row.CreateCell(0);
            ICell cellB1 = row.CreateCell(1);
            row.CreateCell(2).SetCellValue(5.0); // C1
            row.CreateCell(3).SetCellValue(7.0); // D1
            row.CreateCell(4).SetCellValue(9.0); // E1


            cellA1.CellFormula = ("SUM(C1:OFFSET(C1,0,B1))");

            cellB1.SetCellValue(1.0); // range will be C1:D1

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv;
            try
            {
                cv = fe.Evaluate(cellA1);
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("Unexpected ref arg class (NPOI.SS.Formula.LazyAreaEval)"))
                {
                    throw new AssertionException("Identified bug 46948");
                }
                throw e;
            }

            Assert.AreEqual(12.0, cv.NumberValue, 0.0);

            cellB1.SetCellValue(2.0); // range will be C1:E1
            fe.NotifyUpdateCell(cellB1);
            cv = fe.Evaluate(cellA1);
            Assert.AreEqual(21.0, cv.NumberValue, 0.0);

            cellB1.SetCellValue(0.0); // range will be C1:C1
            fe.NotifyUpdateCell(cellB1);
            cv = fe.Evaluate(cellA1);
            Assert.AreEqual(5.0, cv.NumberValue, 0.0);
        }
    }

}